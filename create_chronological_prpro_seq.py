from pyexpat import native_encoding
from typing import Set
import pymiere
from pymiere.wrappers import time_from_seconds
import os
import sys
import win32com.client
from win32com.propsys import propsys, pscon
import pytz
from datetime import datetime
import json
import cv2
import re
import pickle

# To run:
# `python3 ./create_chronological_prpro_seq <relative search path> <sorted files list json filename> <timezone config json filename>`
# Example: `python3 ./create_chronological_prpro_seq .. sorted_files.json timezone_config.json`

# IMPORTANT: Manually import all files into premiere *first*

# ---- PART 1 FUNCTIONS: Sorting the files by earliest date in metadata ---- 

DATE_META = ["Date modified", "Date created", "Date taken", "Date accessed", "Media created"]
VIDEO_EXTENSIONS = [".mp4", ".mov"]

# Add 0's to front of single-digit months and days of the month and hours.
# Assumes string is formatted like '1/2/2022 2:41 PM' for example.
# This would become '01/02/2022 02:41 PM'
# Also handle unicode junk
def clean_date_string(date_string):
    if date_string is None:
        return "NA"
    # remove unicode junk
    date_string = date_string.replace('\u200f', '').replace('\u200e', '')
    # split the date into its parts to see if we need to pad with 0
    date_parts = date_string.split('/')
    # Handle padding month
    if len(date_parts[0]) == 1:
        date_string = '0' + date_string
    # Handle padding day of month
    if len(date_parts[1]) == 1:
        date_string = date_string[:3] + '0' + date_string[3:]
    # add padding to hour if necessary
    if date_string.split(' ')[1].index(':') < 2:
        colon_idx = date_string.index(':')
        date_string = date_string[:colon_idx - 1] + '0' + date_string[colon_idx - 1:]
    return date_string


# https://stackoverflow.com/questions/12521525/reading-metadata-with-python
def get_file_metadata(dir_path, filename):
    # Path shouldn't end with backslash, i.e. "E:\Images\Paris"
    # filename must include extension, i.e. "PID manual.pdf"
    # Returns dictionary containing all file metadata.
    sh = win32com.client.gencache.EnsureDispatch('Shell.Application', 0)
    ns = sh.NameSpace(dir_path)

    file_metadata = {}

    column = 0
    columns = []
    while True:
        colname=ns.GetDetailsOf(None, column)
        if not colname:
            break
        columns.append(colname)
        column += 1

    item = ns.ParseName(str(filename))
    for column in range(len(columns)):
        colval=ns.GetDetailsOf(item, column)
        if colval:
            file_metadata[columns[column]] = colval
    
    for field in DATE_META:
        file_metadata[field] = clean_date_string(file_metadata.get(field))

    if os.path.splitext(filename)[-1].lower() in VIDEO_EXTENSIONS:
        try:
            # https://stackoverflow.com/questions/31507038/python-how-to-read-windows-media-created-date-not-file-creation-date
            properties = propsys.SHGetPropertyStoreFromParsingName(os.path.join(dir_path, filename))
            dt = properties.GetValue(pscon.PKEY_Media_DateEncoded).GetValue()
            file_metadata["Media created"] = dt
        except:
            pass
    return file_metadata

DIMENSIONS_PATTERN = re.compile(r"(\d+) x (\d+)")

# Search the metadata of a file and get its earliest date as well as its dimensions
# because Adobe ExtendScript doesn't have a way to get the dimensions of a photo or video
# from an item that has been imported into Premiere Pro for some reason :/
def get_earliest_date_and_dimensions(filepath, subdir_tz_config):
    
    def determine_timezone(naive_dt):
        utc_dt = pytz.utc.localize(naive_dt)
        # determine which timezone to use by comparing the original time
        # to each timezone config time
        tz_to_use = subdir_tz_config["timezones"][0][1]
        for dt_pair in subdir_tz_config["timezones"][1:]:
            dt = datetime.strptime(dt_pair[0], '%B %d, %Y %I:%M:%S %p')
            # update tz_to_use if the original time is later than a
            # timezone option's time bound in the config
            if utc_dt > pytz.utc.localize(dt):
                tz_to_use = dt_pair[1]
        return tz_to_use

    dir_path = os.path.abspath(os.path.split(filepath)[0])
    filename = os.path.split(filepath)[-1]
    file_meta = get_file_metadata(dir_path, filename)
    correct_dt = None
    # convert to datetime objects for comparison
    try:
        # If the datetime field to use was specified in the config file...
        naive_dt = datetime.strptime(file_meta[subdir_tz_config["datefield"]], '%m/%d/%Y %I:%M %p')
        tz_to_use = determine_timezone(naive_dt)
        correct_dt = naive_dt.astimezone(tz=pytz.timezone(tz_to_use))

    except KeyError: # otherwise just use the earliest datetime
        datetime_meta = []
        for field in DATE_META:
            if file_meta[field] != "NA":
                if isinstance(file_meta[field], datetime):
                    datetime_meta.append(file_meta[field])
                else:
                    # convert to datetime and apply the appropriate timezone
                    # based on what was specified in the timezone config json
                    naive_dt = datetime.strptime(file_meta[field], '%m/%d/%Y %I:%M %p')
                    # determine which timezone to use
                    tz_to_use = determine_timezone(naive_dt)
                    tz_aware_dt = naive_dt.astimezone(tz=pytz.timezone(tz_to_use))
                    datetime_meta.append(tz_aware_dt)
        correct_dt = min(datetime_meta)
    
    
    # Get the dimensions, defaulting to 0 if the dimensions aren't in the metadata and if
    # OpenCV can't calculate them. Some .MOV
    width = 0
    height = 0
    # handle videos
    if os.path.splitext(filepath)[-1].lower() in VIDEO_EXTENSIONS:
        try:
            height = int(file_meta["Frame height"])
            width = int(file_meta["Frame width"])
        except:
            vcap = cv2.VideoCapture(filepath)
            if vcap.isOpened():
                width = int(vcap.get(cv2.CAP_PROP_FRAME_WIDTH))
                height = int(vcap.get(cv2.CAP_PROP_FRAME_HEIGHT))
    # handle photos
    else:
        try:
            height = int(file_meta["Height"])
            width = int(file_meta["Width"])
        except:
            # HEIC files are missing the "Height" and "Width" properties for some reason
            # but they still have "Dimensions"
            try:
                dim_group = re.search(DIMENSIONS_PATTERN, file_meta["Dimensions"])
                width = int(dim_group.group(1))
                height = int(dim_group.group(2))
            except:
                try: # finally, see if OpenCV can get them
                    im = cv2.imread(filepath)
                    height = int(im.shape[0])
                    width = int(im.shape[1])
                except:
                    pass
    if height == 0 or width == 0:
        print("Error getting dimensions of {0}. Dimensions will be left as 0".format(filepath))

    # return a dictionary with the earliest date and dimensions
    return { "filename": filepath,
             "datetime": correct_dt,
             "height": height,
             "width": width }

# Sort all the files based on earliest datetime in metadata
def sort_files(search_root, sorted_json_filename, tz_config):
    print("Retrieving metadata of files...")
    # Get list of all files in windows file explorer (not premiere)
    
    all_files = []
    for (dirpath, _, filenames) in os.walk(search_root):
        if "..\\Auto-Create-Chronological-Premiere-Pro-Sequence" not in dirpath:
            # skip RAW files
            all_files.extend(os.path.join(dirpath, filename) for filename in filenames if os.path.splitext(filename)[1].lower() not in [".cr2", ".cr3"])

    numfiles = len(all_files)

    # Retrieve all the relevant metadata for each file
    filepathnames = set()
    live_photos = []
    file_metas = []
    for i, filepath in enumerate(all_files):
        if i % 100 == 0:
            print("Metadata retrieved from", i, "files of", numfiles)
        print(filepath)
        subdir = os.path.join(*filepath.split('\\')[1:-1])
        file_meta = get_earliest_date_and_dimensions(filepath, tz_config[subdir])
        file_metas.append(file_meta)
        # determine if it's a live photo (if the filepath without the extension has already appeared)
        filepath_wo_ext = os.path.splitext(filepath)[0]
        if filepath_wo_ext in filepathnames:
            live_photos.append((filepath, i))
        else:
            # add to the set of filepathnames we've already seen
            filepathnames.add(os.path.splitext(filepath)[0])
    
    # Handle live photos (it shouldn't be assumed that a live photo's video will come before its image)
    for lp in live_photos:
        # find the photo that corresponds to the live photo's video
        # and assign its datetime to the live photo's video
        dt = next(x["datetime"] for x in file_metas if os.path.splitext(x["filename"])[0] == os.path.splitext(lp[0])[0] and x["filename"] != lp[0])
        file_metas[lp[1]]["datetime"] = dt

    print("Sorting files by datetime...")
    # sort list by earliest date in metadata, which will be found in the 'datetime' field
    sorted_files = sorted(file_metas, key=lambda x: x['datetime'])
    print("Sorted! Saving sorted files metadata in {0}.".format(sorted_json_filename))
    
    # TODO: Sort files with the same datetime by the last four digits of their filename (if applicable)
    # Divide list into sublists based on identical datetimes
    sorted_files_dt_groups = []
    current_dt = None
    for file_meta in sorted_files:
        # create new datetime sublist if it's a new datetime
        if file_meta["datetime"] != current_dt:
            current_dt = file_meta["datetime"]
            sorted_files_dt_groups.append([file_meta])
        else: # otherwise add to the most recent datetime sublist
            sorted_files_dt_groups[-1].append(file_meta)

    # Sort each datetime sublist based on the last 4 digits of name (ex. "IMG_4025.JPG").
    # This won't be applicable for everything but it's better than nothing.
    for dt_group in sorted_files_dt_groups:
        dt_group.sort(key=lambda x: os.path.splitext(x["filename"])[-4])
    
    # Recombine the sorted datetime sublists into the big flat list
    sorted_files = [file_meta for dt_group in sorted_files_dt_groups for file_meta in dt_group]

    # save the list in a JSON file so we hopefully don't have to redo this whole thing again
    with open(sorted_json_filename, "w") as f:
        f.write(json.dumps(sorted_files, default=str, sort_keys=False))

    return sorted_files


# ---- PART 2 FUNCTIONS: Reading the config sequence and generating the new sequence ----

def bin_tree_path_to_filepath(bin_tree_path):
    # Remove the '', project name, and parent bin from the path.
    # example: "\winter tripe.prproj\West Trip Jan 2022\Tim's Photos\IMG_7079.mov"
    # becomes ["Tim's Photos", "IMG_7079.mov"]
    clipProjPath = bin_tree_path.split('\\')[3:]
    # Create the path as it would have been formatted in the JSON file entry for easy lookup.
    # example: ["Tim's Photos", "IMG_7079.mov"] becomes "..\Tim's Photos\IMG_7079.mov"
    return os.path.join(sys.argv[1], *clipProjPath)

# Returns a dictionary representing the bin structure that is
# much faster to search than the bins themselves.
def memoize_bins(parent_bin, pickle_filename):
    print("Creating a dictionary for searching project bins...")
    bin_enum = pymiere.objects.ProjectItemType.BIN
    def memoize_bins_rec(parent_bin, bin_dict, total):
        for child in parent_bin.children:
            if child.type == bin_enum:
                new_dict, total = memoize_bins_rec(child, bin_dict, total)
                bin_dict.update(new_dict)
            else:
                filepath = bin_tree_path_to_filepath(child.treePath)
                bin_dict[filepath] = child
                total += 1
                print("Item {0} catalogued: {1}".format(total, filepath))
        return bin_dict, total
    bin_dict, _ = memoize_bins_rec(parent_bin, {}, 0)
    print("Bin dictionary created! Pickling result in {0} for later!".format(pickle_filename))
    with open(pickle_filename, 'ab') as f:
        pickle.dump(bin_dict, f)
    return bin_dict

# Retreives the info that was stored in the JSON file for a given TrackItem clip.
# Need this because there's no way to obtain the dimensions of a particular photo or video
# via Adobe's API....
def get_clip_filesys_info(clip, sorted_files):
    clip_filepath = bin_tree_path_to_filepath(clip.projectItem.treePath)
    # return the first entry in the sorted list with that filepath/filename
    return next((x for x in sorted_files if x["filename"] == clip_filepath), None)
    
# Reads the configuration sequence in the Premiere Pro project and returns
# a dictionary that specifies what effects and durations should be applied to
# each type of media
def read_config_sequence(project, config_seq_name, sorted_files):
    config_seq = next((x for x in project.sequences if x.name == config_seq_name), None)
    if not config_seq:
        print("ERROR: No sequence named {0} has been found in the open Premiere Pro project!".format(config_seq_name))
        exit(1)
    
    prop_dict = { "photo": {}, "video": {} }

    # Get the necessary properties for each clip in the first video track.
    print("Reading {0} to learn what to do with each type of photo and video...".format(config_seq_name))
    for clip in config_seq.videoTracks[0].clips:
        print("Analyzing {0}...".format(clip.name))
        # get info with video dimensions:
        clipInfo = get_clip_filesys_info(clip, sorted_files)
        if clipInfo is None:
            print(("CONFIG SEQUENCE WARNING: "
                   "Dimensions for {0} could not be found. "
                   "Its effect properties will not be recorded.").format(clip.projectItem.treePath))
        else:
            height = clipInfo["height"]
            width = clipInfo["width"]
            # get the "Scale" property of the "Motion" component
            motion = next(x for x in clip.components if x.displayName == "Motion")
            scale = next(x for x in motion.properties if x.displayName == "Scale")
            
            # if it's a video...
            if os.path.splitext(clip.name)[-1].lower() in VIDEO_EXTENSIONS:
                # assume that this lacks keyframes and is the only property we care about
                prop_dict["video"][(height, width)] = { "scale": scale.getValue() }
            
            # else assume it's a photo
            else:
                # for photos we care about the clip duration
                prop_dict["photo"][(height, width)] = { "duration": clip.duration }
                # we also care about the Scale keyframes
                # assume that there are 2 keyframes, 1 at beginning and 1 at end of clip
                prop_dict["photo"][(height, width)]["scaleInKey"] = scale.getValueAtTime(clip.inPoint)
                prop_dict["photo"][(height, width)]["scaleOutKey"] = scale.getValueAtTime(clip.outPoint)
                # We care about Position keyframes too but those will just be added by default
                # assuming they are co-located with the Scale keyframes but they will retain their default values.
                # So no need to record anything for them here (for this project at least).
    
    print("Finished reading {0}!".format(config_seq_name))
    return config_seq.getSettings(), prop_dict

# Some clips may not exactly match the dimensions of those that were
# used in the configuration sequence, so here we'll find the closest one.
# For now just go based on the height... usually more important than width
def calculate_closest_dimensions(file_info, media_dict):
    return sorted([x for x in media_dict.keys()], key=lambda y: abs(file_info["height"] - y[0]))[0]

# Populate the new sequence with the photos and videos in the correct order
# with the correct motion properties applied
def add_clips_to_sequence(new_seq, sorted_files, prop_dict, bin_dict):
    print("Adding clips to sequence in chronological order...")
    track = new_seq.videoTracks[0]
    seq_time = pymiere.Time()
    seq_time.seconds = 0
    # get the time duration of a single frame
    frameTime = pymiere.Time()
    frameTime.ticks = str(new_seq.timebase)
    num_files = len(sorted_files)
    for i, file_info in enumerate(sorted_files):
        proj_item = None
        # there shouldn't be any RAW files but skip them to remain sane anyways
        if os.path.splitext(file_info['filename'])[-1].lower() in [".cr2", ".cr3"]:
            print("{0} of {1}: Skipping RAW file {2}".format(str(i + 1), num_files, file_info['filename']))
            continue
        try:
            print("{0} of {1}: Adding {2} to the sequence and applying motion properties...".format(str(i+1), num_files, file_info["filename"]))
            proj_item = bin_dict[file_info["filename"]]
        except KeyError:
            print(("WARNING: {0} appears to be missing from the Premiere project files "
                       "and will be skipped!").format(file_info['filename']))
            continue
        # add the projectItem to the sequence
        track.insertClip(proj_item, seq_time.seconds)
        # apply the appropriate Motion properties
        new_clip = track.clips[i]
        motion = next(x for x in new_clip.components if x.displayName == "Motion")
        scale = next(x for x in motion.properties if x.displayName == "Scale")
        # video
        if os.path.splitext(file_info['filename'])[-1].lower() in VIDEO_EXTENSIONS:
            dimensions = calculate_closest_dimensions(file_info, prop_dict["video"])
            scale.setValue(prop_dict["video"][dimensions]["scale"], True)
        # photo
        else:
            position = next(x for x in motion.properties if x.displayName == "Position")
            dimensions = calculate_closest_dimensions(file_info, prop_dict["photo"])
            new_clip.end = time_from_seconds(seq_time.seconds + prop_dict["photo"][dimensions]["duration"].seconds)
            scale.setTimeVarying(True)
            position.setTimeVarying(True)
            # start keyframes
            scale.addKey(new_clip.inPoint.seconds)
            scale.setValueAtKey(new_clip.inPoint.seconds, prop_dict["photo"][dimensions]["scaleInKey"], 1)
            position.addKey(new_clip.inPoint.seconds)
            position.setValueAtKey(new_clip.inPoint.seconds, [0.5, 0.5], 1)
            # end keyframes
            # I like to have each clip's end keyframes occur at the start of
            # the last frame for which the clip is visible.
            outTime = new_clip.inPoint.seconds + new_clip.duration.seconds - frameTime.seconds
            print("outTime:", outTime)
            scale.addKey(outTime)
            scale.setValueAtKey(outTime, prop_dict["photo"][dimensions]["scaleOutKey"], 1)
            position.addKey(outTime)
            position.setValueAtKey(outTime, [0.5, 0.5], 1)
        seq_time.seconds += new_clip.duration.seconds
        
        

# ------------------------------------------------

def main():
    # ---- PART 1: Organize the files ----
    search_root = sys.argv[1]
    sorted_json_filename = sys.argv[2]
    tz_config_filename = sys.argv[3]

    project = pymiere.objects.app.project

    # Before doing anything else, verify that the specified directory is present as a bin
    # in the Premiere Pro project.
    root_dirname = os.path.basename(os.path.abspath(search_root))
    parent_bin = next((x for x in project.rootItem.children if x.name == root_dirname), None)
    if not parent_bin:
        print(("ERROR: Bin \"{0}\" not found.\n"
               "There should be a bin named {0} in the root of the project explorer"
               "in Premiere that contains the subdirectories with the media!").format(root_dirname))
        exit(1)

    # Obtain the information on timezones
    with open(tz_config_filename, "r", encoding="utf-8") as f:
        tz_config = json.load(f)
        print("Loaded timezone information from {0}".format(tz_config_filename))

    # only get all the metadata and sort it if we haven't done that before
    sorted_files = []
    if os.path.exists(sorted_json_filename):
        with open(sorted_json_filename, "r") as f:
            sorted_files = json.load(f)
        print("Sorted files relevant metadata loaded from {0}".format(sorted_json_filename))
    else:
        sorted_files = sort_files(search_root, sorted_json_filename, tz_config)
    
    # ---- PART 2: Read the Premiere Pro config sequence and generate the new sequence
    
    # First memoize the contents of the bins in the Premiere Pro file
    # because Premiere is incredibly slow at searching the bins....
    # It's better to just do it once.
    pickle_filename = "bin_dict_pkl"
    if os.path.exists(pickle_filename):
        with open(pickle_filename, 'rb') as f:
            bin_dict = pickle.load(f)
        print("Loaded previously pickled project bin dictionary from {0}".format(pickle_filename))
    else:
        bin_dict = memoize_bins(parent_bin, pickle_filename)

    # Next read the "config_sequence" to decide how to handle each media type
    seq_settings, prop_dict = read_config_sequence(project, "config_sequence", sorted_files)

    # create a new sequence using the same settings as the config sequence
    new_seq_name = "GENERATED_SEQUENCE"
    print("Go click \"OK\" on the alert in Premiere!")
    pymiere.objects.alert(("You're about to be prompted to create a new sequence.\n"
                           "Just click \"OK\" and don't worry about it!"))
    print("Creating new sequence {0}...".format(new_seq_name))
    project.createNewSequence(new_seq_name, "placeholderID")
    # user might need to click "Okay" in Premiere here
    new_seq = next(x for x in project.sequences if x.name == new_seq_name)
    new_seq.setSettings(seq_settings)
    print("Sequence {0} has been created!".format(new_seq_name))

    # Populate the new sequence with the photos and videos in the correct order
    # with the correct motion properties applied
    add_clips_to_sequence(new_seq, sorted_files, prop_dict, bin_dict)

    print("Finished adding all {0} clips to {1}!".format(len(sorted_files), new_seq_name))
    exit(0)

if __name__ == "__main__":
    main()

