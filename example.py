#!/usr/bin/env python

"""
Program:
This program will record log into Excel.
History:
20170707 Kuanlin Chen
"""

import xlwt
import sys
import subprocess
import string
import codecs

filename = "example.xls"
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet1")

def main(orig_args):
	output(filename)

def output(filename):
	
	#proc = subprocess.Popen('ls -l',shell=True,stdout=subprocess.PIPE)
	proc = subprocess.Popen(['ls','-l'],stdout=subprocess.PIPE)
	text = proc.stdout.read()
	#print(text)

	text = text.split('\n')
	#print(text)
	i = 0
	for line in text:
		#split() method without any argument splits on whitespace
		line = line.split()
		print(line)
		j = 0
		for word in line:
			sheet1.write(i,j,word)
			j = j+1

		i = i+1

	book.save(filename)

if __name__ == '__main__':
    main(sys.argv)