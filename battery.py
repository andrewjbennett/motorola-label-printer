import win32api
import win32print
import tempfile
import subprocess
import time
import os
import sys
import csv
from itertools import izip

# a script to take input from the Motorola IMPRES battery reader
# and use this to generate a printed label containing essential
# information about the battery.

# this has been tested to run on windows xp [yes, I know].
# in order to make it work, the battery printer must be set as
# the default printer, and the text size of Notepad must be
# set to the size you'd like for it to fit nicely on the label.


# run the battery reader software, which automatically logs to a file
process = subprocess.Popen(["C:\\Program Files\Motorola Solutions\IMPRES Battery Reader\BatteryReader.exe"])

# wait a while to make sure the process runs
time.sleep(4)


# where the file is set from the program to be stored
filename = "C:/Python27/battery.csv"

# if we have an error, we want to print something later on
error = False



# this is literally the worst, I feel vey uncomfortable publishing
# this code with this in it.

# the file format ended up being rotated to the original file, and so
# this will "pivot" the data, and write it to another file in the
# "correct" orientation.

a = izip(*csv.reader(open(filename, "rb")))
csv.writer(open("output.csv", "wb")).writerows(a)


# open the csv file, and read it into a csv reader
lines = open("output.csv").readlines()
r = csv.reader(lines)


text_out = ""

done = 0

# we store these separately because they need to be
# processed later on differently
date = ""
recommendations = ""
num = ""

# go through and check for certain lines
for line in r:
    if "Kit Number" in line:
        num = line[1]
    if "Log Date" in line:
        date += "%s: %s" % (line[0].split(" ")[1], line[1][1:])
    if "Present Charge" in line:
        text_out += "%s: %s mAh\n" % (line[0], line[1])
    if "% of Rated Capacity" in line:
        text_out += "%s: %s %%\n" % (line[0], line[1])
    if "Voltage" in line:
        text_out += "%s: %s V\n" % (line[0], line[1])
    if "Recommendations" in line:
        recommendations += "%s: %s\n" % (line[0], (line[1].split(";")[0]))
        if any("Error" in s for s in line):
            error = True
        
# process the output to have what we want in it
text_out = date + " | " + num + "\n" + text_out 

text_out += recommendations

if error:
    text_to_print = "Warning: REPLACE BATTERY\n\n"
else: 
    text_to_print = ""

text_to_print += text_out

# removed for privacy purposes
text_to_print += "COMPANY NAME | PHONE NUMBER"

# remove the battery log file
os.remove(filename)

# print the label
tmpfile = tempfile.mktemp(".txt")
open(tmpfile, "w").write(text_to_print)
win32api.ShellExecute(0, "print", tmpfile, None, ".", 0)

