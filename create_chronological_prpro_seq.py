import pymiere
import os
import sys
import win32com.client
from datetime import datetime
import json

# To run:
# `python3 ./create_chronological_prpro_seq <relative search path>` `<sorted files list json filename>`
# Example: `python3 ./create_chronological_prpro_seq .. sorted_files.json`

# IMPORTANT: Manually import all files into premiere *first*

# ---- PART 1 FUNCTIONS: Sorting the files by earliest date in metadata ---- 

DATE_META = ["Date modified", "Date created", "Date taken", "Date accessed", "Media created"]

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
    
# Search the metadata of a file and get its earliest date
def get_earliest_date(filepath):
    dir_path = os.path.abspath(os.path.split(filepath)[0])
    filename = os.path.split(filepath)[-1]
    file_meta = get_file_metadata(dir_path, filename)
    # convert to datetime objects for comparison
    datetime_meta = []
    for field in DATE_META:
        if file_meta[field] != "NA":
            datetime_meta.append(datetime.strptime(file_meta[field], '%m/%d/%Y %I:%M %p'))
    # return earliest date
    return min(datetime_meta)

# Sort all the files based on earliest datetime in metadata
def sort_files(search_root, sorted_json_filename):
    # Get list of all files in windows file explorer (not premiere)
    
    all_files = []
    for (dirpath, _, filenames) in os.walk(search_root):
        if dirpath != "..\\Premiere Pro script":
            all_files.extend(os.path.join(dirpath, filename) for filename in filenames)

    numfiles = len(all_files)

    datetime_lst = []
    for i, f in enumerate(all_files):
        if i % 100 == 0:
            print("Metadata retrieved from", i, "files of", numfiles)
        print(f)
        datetime_lst.append(get_earliest_date(f))

    print("Sorting files by datetime...")
    # sort list by earliest date in metadata
    sorted_files = [x for _, x in sorted(zip(datetime_lst, all_files))]
    print("Sorted!")

    # save the list in a JSON file so we hopefully don't have to redo this whole thing again
    with open(sorted_json_filename, "w") as f:
        f.write(json.dumps(sorted_files))

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
    print(config_seq.name)


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
    
    # ---- PART 2: Create the Premiere Pro sequence

    project = pymiere.objects.app.project

    # First read the "config_sequence" to decide how to handle each media type
    read_config_sequence(project, "config_sequence")
    

# find the corresponding file by name in project bins

# for each imported file, use track.insertClip() to insert it into the sequence

if __name__ == "__main__":
    main()

