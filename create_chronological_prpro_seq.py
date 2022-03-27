from turtle import width
import pymiere
import os
import sys
import win32com.client
from datetime import datetime
import json
import cv2
import re

# To run:
# `python3 ./create_chronological_prpro_seq <relative search path> <sorted files list json filename>`
# Example: `python3 ./create_chronological_prpro_seq .. sorted_files.json`

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

    colnum = 0
    columns = []
    while True:
        colname=ns.GetDetailsOf(None, colnum)
        if not colname:
            break
        columns.append(colname)
        colnum += 1

    item = ns.ParseName(str(filename))
    for colnum in range(len(columns)):
        colval=ns.GetDetailsOf(item, colnum)
        if colval:
            file_metadata[columns[colnum]] = colval
    
    for field in DATE_META:
        file_metadata[field] = clean_date_string(file_metadata.get(field))

    return file_metadata

DIMENSIONS_PATTERN = re.compile(r"(\d+) x (\d+)")

# Search the metadata of a file and get its earliest date as well as its dimensions
# because Adobe ExtendScript doesn't have a way to get the dimensions of a photo or video
# from an item that has been imported into Premiere Pro for some reason :/
def get_earliest_date_and_dimensions(filepath):
    dir_path = os.path.abspath(os.path.split(filepath)[0])
    filename = os.path.split(filepath)[-1]
    file_meta = get_file_metadata(dir_path, filename)
    # convert to datetime objects for comparison
    datetime_meta = []
    for field in DATE_META:
        if file_meta[field] != "NA":
            datetime_meta.append(datetime.strptime(file_meta[field], '%m/%d/%Y %I:%M %p'))
    
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
             "datetime": min(datetime_meta),
             "height": height,
             "width": width }

# Sort all the files based on earliest datetime in metadata
def sort_files(search_root, sorted_json_filename):
    # Get list of all files in windows file explorer (not premiere)
    
    all_files = []
    for (dirpath, _, filenames) in os.walk(search_root):
        if "..\\Auto-Create-Chronological-Premiere-Pro-Sequence" not in dirpath:
            all_files.extend(os.path.join(dirpath, filename) for filename in filenames)

    numfiles = len(all_files)

    file_metas = []
    for i, filepath in enumerate(all_files):
        if i % 100 == 0:
            print("Metadata retrieved from", i, "files of", numfiles)
        print(filepath)
        file_meta = get_earliest_date_and_dimensions(filepath)
        file_metas.append(file_meta)

    print("Sorting files by datetime...")
    # sort list by earliest date in metadata
    sorted_files = sorted(file_metas, key=lambda x: x['datetime'])
    print("Sorted!")

    # save the list in a JSON file so we hopefully don't have to redo this whole thing again
    with open(sorted_json_filename, "w") as f:
        f.write(json.dumps(sorted_files, default=str, sort_keys=False))

    return sorted_files


# ---- PART 2 FUNCTIONS: Reading the config sequence ----

def read_config_sequence(project, config_seq_name):
    config_seq = None
    for seq in project.sequences:
        if seq.name == config_seq_name:
            config_seq = seq
            break
    
    if not config_seq:
        print("ERROR: No sequence named {0} has been found in the open Premiere Pro project!".format(config_seq_name))
        exit(1)
    
    prop_dict = { "photo": {}, "video": {} }

    # Get the necessary properties for each clip in the first video track.
    for clip in config_seq.videoTracks[0].clips:
        # if it's a video...
        if os.path.splitext(clip.name)[-1].lower() in VIDEO_EXTENSIONS:
            # TODO record video dimensions
            
            # TODO record video Motion property
            for component in clip.components:
                if component.displayName == "Motion":

                    break

        # else assume it's a photo
        else:
            continue


    return config_seq.getSettings(), prop_dict





# ------------------------------------------------

def main():
    # ---- PART 1: Organize the files ----
    search_root = sys.argv[1]
    sorted_json_filename = sys.argv[2]
    # only get all the metadata and sort it if we haven't done that before
    sorted_files = []
    if os.path.exists(sorted_json_filename):
        with open(sorted_json_filename, "r") as f:
            sorted_files = json.load(f)
    else:
        sorted_files = sort_files(search_root, sorted_json_filename)
    
    # ---- PART 2: Read the Premiere Pro config sequence

    project = pymiere.objects.app.project

    # First read the "config_sequence" to decide how to handle each media type
    read_config_sequence(project, "config_sequence")
    

# find the corresponding file by name in project bins

# for each imported file, use track.insertClip() to insert it into the sequence

if __name__ == "__main__":
    main()

