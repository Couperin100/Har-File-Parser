# Har-File-Parser
A script to parse a Har file and export it into an Excel spreadsheet

This script uses the xlwt library to export the data into Excel but the rest of the libraries are standard.  It was created to help application support guys grab the http requests from a customers machine in order to troubleshoot later if the call ran over time.  The HLS sheet would contain all HLS segments from any of the requests in the main sheet.  This is an added bonus if you're trying to troubleshoot HLS streaming requests in regards to the speed of each segment. POST, GET, DELETES are all colour coded for easier reading and indentifing.

## Installing

1. Create a new directory and save this script inside it.
2. make sure you have installed xlwt (https://pypi.python.org/pypi/xlwt)

## How to use

To use you simply need a Har file taken from Google Chrome by viewing a webpage with the dev tools up, right clicking and then selecting "save to har file" from the list.  Once done save it in the same folder as this script and run it like the below:

python har_file.py my_saved_har_file.har

It'll create a spreadsheet and save it in the same directory.  
