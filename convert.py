#!/usr/bin/env python

__author__ = "Quique Molla <emolladev@gmail.com>"
__version__ = "0.0.1"
__license__ = "GPL-2+"

import os
import json
import sys
import ast

try:
  from argparse import ArgumentParser
except ImportError:
  print("\n******************************************************************")
  print("    argparse is required to run this script.")
  print("    It is available here: https://pypi.python.org/pypi/argparse")
  print("****************************************************************\n")
  sys.exit(1)
try:
  import xlrd
except ImportError:
  print("\n******************************************************************")
  print("    xlrd is required to run this script.")
  print("    It is available here: http://pypi.python.org/pypi/xlrd")
  print("****************************************************************\n")
  sys.exit(1)
try:
  import csv
except ImportError:
  print("\n******************************************************************")
  print("    csv is required to run this script.")
  print("    It is available here: http://pypi.python.org/pypi/csv")
  print("****************************************************************\n")
  sys.exit(1)

try:
  import chardet
except ImportError:
  print("\n******************************************************************")
  print("    csv is required to run this script.")
  print("    It is available here: http://pypi.python.org/pypi/chardet")
  print("****************************************************************\n")
  sys.exit(1)

def getFileType(l, headersFile, encoding="utf8"):
    row = l.lower().replace(' ', '')
    filetype = ""

    with open(headersFile) as json_data:
        headerArray = json.load(json_data)
        for value in headerArray:
            allin = True
            for field in headerArray[value]:
                fd = field.lower().replace(' ', '')
                if fd not in row.decode(encoding):
                    allin = False
            if allin:
                filetype = value

    return filetype, row


def read_excel_xlrd(inputFilePath, outFilePath, headerFields):
    wb = xlrd.open_workbook(inputFilePath) #on_demand = True, encoding='cp1252'
    outfile = open(outFilePath, "w")
    filetype = ""
    numlines = 0
    for sheet_name in wb.sheet_names():
        sh = wb.sheet_by_name(sheet_name)
        if sh.ncols == 1:
            sniffed = csv.Sniffer().sniff(''.join(sh.row_values(0)))
            for i in range(sh.nrows):
                row = ''.join(map(lambda e: unicode(e).strip(), ';'.join(sh.row_values(i)).split(sniffed.delimiter))) + '\n'

                if filetype == "":
                    filetype, header = getFileType(row, headerFields)
                    if header:
                        outfile.write(header)
                else:
                    outfile.write(row.encode("utf8"))
                    numlines = numlines + 1
        else:
            for i in range(sh.nrows):
                row = ';'.join(map(lambda e: unicode(e).strip(), sh.row_values(i))) + '\n'

                if filetype == "":
                    filetype, header = getFileType(row, headerFields)
                    if header:
                        outfile.write(header)
                else:
                    outfile.write(row.encode("utf8"))
                    numlines = numlines + 1
    outfile.close()
    return filetype, numlines


def read_csv(inputFilePath, outFilePath, headerFields):
    outfile = open(outFilePath, 'w')
    filetype = ""
    numlines = 0
    with open(inputFilePath, 'rU') as file:
        sniffed = csv.Sniffer().sniff(file.readline())
        detectedEncoding = chardet.detect(file.readline())
        encoding = "utf-8"
        if detectedEncoding is not None:
            encoding = detectedEncoding['encoding']
        while sniffed.delimiter not in ";,\t":
            sniffed = csv.Sniffer().sniff(file.readline())
        file.seek(0)
        reader = csv.reader(file, delimiter=sniffed.delimiter)
        for i in reader:
            row = ';'.join(map(lambda e: str(e).strip(), i)) + '\n'
            if filetype == "":
                row = ';'.join(map(lambda e: str(e).strip(), i)) + '\n'
                filetype, header = getFileType(row, headerFields, encoding)
                if header:
                    outfile.write(header)
            else:
                outfile.write(row)
                numlines = numlines + 1
    outfile.close()
    return filetype, numlines


def filter_to_csv(inputFilePath, extension, outputFilePath, headerFields):
    filetype = ""
    numlines = 0

    try:
        if inputFilePath.lower().endswith(".csv") or extension == "csv":
            filetype, numlines = read_csv(inputFilePath, outputFilePath, headerFields)
        else:
            filetype, numlines = read_excel_xlrd(inputFilePath, outputFilePath, headerFields)
    except Exception:
        filetype, numlines = read_csv(inputFilePath, outputFilePath, headerFields)

    return filetype, numlines


if __name__ == "__main__":
  parser = ArgumentParser(usage="%%prog -i infile -o outfile -hf headersFile")
  parser.add_argument('-i', "--infile", dest="inputFile", required=True, help="full path to file of the source spreadsheet")
  parser.add_argument('-e', "--extension", dest="extension", help="file extension")
  parser.add_argument('-o', "--outfile", dest="outputFile", help="filename to save the generated csv file")
  parser.add_argument('-hf', "--headersFile", dest="headersFile", help="header fields to detect in input file")

  args = parser.parse_args()
  if not len(sys.argv) > 3:
    parser.print_help()
  else:
      if not os.path.exists(args.inputFile):
        parser.error("File " + os.path.abspath(args.inputFile) + " not existss")
      if not os.path.exists(args.headersFile):
        parser.error("File " + os.path.abspath(args.headersFile) + " not exists")

      if not os.path.isfile(args.inputFile):
        parser.error("File " + os.path.abspath(args.inputFile) + " is a folder")
      if os.path.exists(args.outputFile) and not os.path.isfile(args.outputFile):
          basename = os.path.basename(args.inputFile)
          args.outputFile = args.outputFile + os.path.sep + os.path.splitext(basename)[0]
      if not os.path.isfile(args.headersFile):
        parser.error("File " + os.path.abspath(args.headersFile) + " is a folder")
      if not args.outputFile or args.outputFile == "undefined":
        args.outputFile = args.inputFile + ".csv"

      filetype, numLines = filter_to_csv(args.inputFile, args.extension, args.outputFile, args.headersFile)
      sys.stdout.write(json.JSONEncoder().encode({"fileType":filetype,"numLines":numLines,"outFile":os.path.abspath(args.outputFile)}))
