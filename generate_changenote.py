#!/usr/bin/env python
#Program:
#This program will write change note into Excel.
#History:
#20170706 Kuanlin Chen

import subprocess
import sys
import string
import re
import xlwt

filename = "changenote.xls"
fixversion = "V08d"
checkresult = "OK"

book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

#sincedate = "2017-06-25"
#command = 'repo forall -c \'git log --oneline --date=format:"%Y-%m-%d" --pretty="<<<$REPO_PROJECT<<<%s\n\n%b<<<%cd<<<%an>>>" --name-only --since="2017-05-25"\''

def main(orig_args):
    sincedate = raw_input("Since which Date? (yyyy-mm-dd) ")
    #if sincedate is empty, repeat raw_input
    while not sincedate:
        sincedate = raw_input("Since which Date? (yyyy-mm-dd) ")

    print("Create ChangeNote between "+sincedate+" and Today")
    output(filename,sincedate)


def output(filename,sincedate):
    #Set up column size
    default_width = 256 * 20

    sheet1.write(0,0,'Defect Description')
    sheet1.write(0,1,'Developer')
    sheet1.write(0,2,'Fixed on Date')
    sheet1.write(0,3,'Changed File')
    sheet1.write(0,4,'TD Issue')
    sheet1.write(0,5,'Detected')
    sheet1.write(0,6,'Fixed in Version')
    sheet1.write(0,7,'Check Result')
    sheet1.row(0).height_mismatch = True
    sheet1.row(0).height = 256*1

    proc = subprocess.Popen('repo forall -c \'git log --oneline --no-merges --date=format:"%Y-%m-%d" --pretty="<<%cd<<%an<<%s\n\n%b>>" --name-only --since='+sincedate+'\'',shell=True, stdout=subprocess.PIPE)
    text = proc.stdout.read()

    #Split output by >>
    #Notice that changed file will be output in the last space which after >>>.
    #So we need to modify the place of changed file later.
    s = string.split(text,'\n>>')
    i = 1
    for line in s:
        j = 3
        innerline = string.split(line,'<<')
        sheet1.row(i).height_mismatch = True
        sheet1.row(i).height = 256*2
        for line in innerline:
            line = line.lstrip()
            line = line.rstrip('\n')
            if j==0:
                #Set the first cloumn bigger.
                col_width = 256 * 40
                sheet1.col(j).width = col_width
                tdissue(line,i)
            else:
                #Set up column size
                sheet1.col(j).width = default_width

            if j==3:
                #Move changed file to Previous colume
                tmp = i-1
                if tmp!=0:
                    sheet1.write(tmp,j,line)
            else:
                sheet1.write(i,j,line)

            j = j-1
        versionandresult(i)
        i = i+1

    book.save(filename)


def versionandresult(i):
    #Set up the default value of fixversion and checkresult.
    cell = i-1
    if cell!=0:
        sheet1.write(cell,6,fixversion)
        sheet1.write(cell,7,checkresult)


def tdissue(line,i):
    #Search TD name and write to column.
    if "TD" in line:
        mylist = re.split("]",line)
        tdnumber = mylist[5]
        tdnumber = tdnumber.lstrip('[')
        print(tdnumber)
        sheet1.write(i,4,tdnumber)
        sheet1.write(i,5,'TD')
    else:
        print("No")
        sheet1.write(i,4,'-')
        sheet1.write(i,5,'RD')


if __name__ == '__main__':
    main(sys.argv)
