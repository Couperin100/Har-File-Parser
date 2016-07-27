import json
import sys
import xlwt
from urlparse import urlparse


def har_parser(harfile_path):

    headings = ["Entry Number","URL","Domain","Size","Time","Mimetype"]
    book = xlwt.Workbook()
    ws = book.add_sheet("HTTP Requests", cell_overwrite_ok=False)
    ws_hls = book.add_sheet("HLS Segments", cell_overwrite_ok=False)


    harfile = open(harfile_path)
    harfile_json = json.loads(harfile.read())
    i = 0
    data = []
    data1 = []

    for http_entry in harfile_json['log']['entries']:

        method = http_entry['request']['method']
        response = http_entry['response']['status']
        url = http_entry['request']['url']
        urlname = urlparse(http_entry['request']['url'])
        size_bytes = http_entry['response']['bodySize']
        time_taken = http_entry['time']
        timings = http_entry['timings']
        useragent = http_entry['request']['headers']

        for item in useragent:
            if "User-Agent" in item['name']:
                useragent1 = item['value']

        hls = urlparse(http_entry['request']['url'])
        hls_list = [i,hls.params,timings['blocked'],timings['dns'],timings['connect'],timings['send'], timings['wait'],timings['receive'], timings['ssl'],time_taken]
        if "seg" in hls.params:
            data1.append(hls_list)

        mimetype = 'unknown'
        if 'mimeType' in http_entry['response']['content']:
            mimetype = http_entry['response']['content']['mimeType']
        
        mylist=[i,method,response,url,urlname.hostname,size_bytes,time_taken,mimetype,useragent1]
        data.append(mylist)
        i = i+1

    style = xlwt.XFStyle()
    style1 = xlwt.XFStyle()
    style2 = xlwt.XFStyle()
    style.alignment.wrap = 1
    style.alignment.vert = xlwt.Alignment.VERT_TOP
    style1.alignment.wrap = 1
    style1.alignment.vert = xlwt.Alignment.VERT_TOP
    style2.alignment.wrap = 1
    style2.alignment.vert = xlwt.Alignment.VERT_TOP
    pattern = xlwt.Pattern()
    pattern.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern1 = xlwt.Pattern()
    pattern1.pattern = xlwt.Pattern.SOLID_PATTERN
    pattern2 = xlwt.Pattern()
    pattern2.pattern = xlwt.Pattern.SOLID_PATTERN
    style.pattern = pattern
    style1.pattern = pattern1
    style2.pattern = pattern2

    ws.write(2,0, "Entry")
    ws.write(2,1, "Status")
    ws.write(2,2, "Code")
    ws.write(2,3, "URL")
    ws.write(2,4, "Host")
    ws.write(2,5, "Size")
    ws.write(2,6, "Time")
    ws.write(2,7, "Mimetype")
    ws.write(2,8, "User-Agent")

    ws_hls.write(2,0, "Entry")
    ws_hls.write(2,1, "Segment")
    ws_hls.write(2,2, "Blocked")
    ws_hls.write(2,3, "DNS")
    ws_hls.write(2,4, "Connect")
    ws_hls.write(2,5, "Send")
    ws_hls.write(2,6, "Wait")
    ws_hls.write(2,7, "Receive")
    ws_hls.write(2,8, "SSL")
    ws_hls.write(2,9, "Total Time Taken (ms)")

    for rowx,row in enumerate(data, 3):
        for colx, value in enumerate(row):
            if "POST" in row[1]:
                pattern.pattern_fore_colour = 5
                ws.write(rowx, colx, value, style)
            elif "DELETE" in row[1]:
                pattern2.pattern_fore_colour = 2
                ws.write(rowx, colx, value, style2)
            else:
                pattern1.pattern_fore_colour = 3
                ws.write(rowx, colx, value, style1)
    
    for rowx, row in enumerate(data1, 3):
        for colx, value in enumerate(row):
            ws_hls.write(rowx, colx, value)

    ws.col(0).width = 1400
    ws.col(1).width = 2000
    ws.col(2).width = 1400
    ws.col(3).width = 44000
    ws.col(4).width = 6400
    ws.col(5).width = 3400
    ws.col(6).width = 6400
    ws.col(7).width = 6400
    ws.col(8).width = 16400

    ws_hls.col(0).width = 1400
    ws_hls.col(1).width = 3400
    ws_hls.col(2).width = 3400
    ws_hls.col(3).width = 3400
    ws_hls.col(4).width = 3400
    ws_hls.col(5).width = 3400
    ws_hls.col(6).width = 3400
    ws_hls.col(7).width = 3400
    ws_hls.col(8).width = 3400
    ws_hls.col(9).width = 6400

    book.save("HAR_spreadsheet.xls")
    
har_parser(sys.argv[1])
