# Har-File-Parser
A script to parse a Har file and export it into an Excel spreadsheet

This script uses the xlwt library to export the data into Excel but the rest of the libraries are standard.  It was created to help application support guys grab the http requests from a customers machine in order to troubleshoot later in the call ran over time.  The HLS sheet would contain all HLS segments from any of the requests in the main sheet.  This is an added bonus if you're trying to troubleshoot HLS streaming requests in regards to the speed of each segment.

## How to use

