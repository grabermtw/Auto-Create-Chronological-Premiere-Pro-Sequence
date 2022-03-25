# Auto-Create-Chronological-Premiere-Pro-Sequence

## Overview
This is a script for creating a sequence in Premiere Pro containing all the media in a directory arranged by earliest datetime present in the file metadata. It will search `Date modified`, `Date created`, `Date taken`, `Date accessed`, and `Media created` and use the earliest date available from those for sorting to yield the greatest sorting accuracy possible (at least for the case I'm using it for).

I'm creating this because Premiere Pro itself was for some reason unable to correctly sort the media by date (for me at least). Might've been just something I was doing wrong, but regardless, writing this script is the easiest solution I've found.

This was written on a Windows 10 machine, and almost certainly wouldn't work on Mac or Linux without modification of Part 1 (where the directories are searched and the media is sorted), as that uses `win32com.client` for getting the metadata of each file. But if anyone wants to attempt that modification, be my guest!

## Use

### Premire Pro Preparation
First, Premiere Pro needs to be running with the project you want to add the sequence to already open. You should also have already imported all of the media that you want to be sorted and placed in the sequence, and that imported media should be organized the same way inside Premiere Pro as it is organized in your filesystem. Basically, if you run the script with the search path pointing to a parent directory containing subdirectories each containing photos and videos, then there should be a parent directory (or "bin" as Adobe likes to call it) inside the Premiere Pro project with the same subdirectories/bins as was on the filesystem. **This script will _not_ import the files into your project for you.** That would just take too long for big projects (or at least the one I'm creating this for).

Next, create a sequence in your Premiere Pro project called "config_sequence", and add an example of every media type that the script will encounter, and place any motion effects on each one that you want the script to apply to each instance of that media type it encounters. The script will read that sequence and create a new sequence with identical properties, and then add the media in chronological order, deciding what to do with each photo or video based on what was done with the photo or video most similar to it in the config_sequence. It will determine this based on both media type (photo vs video) and the dimensions of the media. It will _not_ consider the file extension (i.e. it will handle a JPG and PNG identically if they have the same dimensions).

### Running the script
To run: `python3 ./create_chronological_prpro_seq <relative search path> <sorted files list json filename>`
Example: `python3 ./create_chronological_prpro_seq .. sorted_files.json`

The `<sorted files list json filename>` will be created if it doesn't already exist, and it will contain a sorted list of all the media. If it already exists, then the script will read from this file rather than going through the sorting process again, as the sorting process can take hours in some cases depending on the media types it is sorting.


I'd like to thank [Pymiere](https://github.com/qmasingarbe/pymiere) for making this possible, and I'd like to ask Adobe why their absurdly expensive software-as-a-service can't seem to provide this service itself.
