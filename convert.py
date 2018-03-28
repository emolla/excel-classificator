#!/usr/bin/env python

__author__ = "Quique Molla <emolladev@gmail.com>"
__version__ = "0.0.1"
__license__ = "GPL-2+"

import os
import io
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
  # import csv
  from backports import csv
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

    with io.open(headersFile, "r", encoding=encoding) as json_data:
        headerArray = json.load(json_data)
        for value in headerArray:
            allin = True
            for field in headerArray[value]:
                fd = field.lower().replace(' ', '')
                if fd not in row:
                    allin = False
            if allin:
                filetype = value

    return filetype, row


def read_excel_xlrd(inputFilePath, outFilePath, headerFields, detect):
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
                    if detect == "true":
                        break
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
                    if detect == "true":
                        break
                    outfile.write(row.encode("utf8"))
                    numlines = numlines + 1
    outfile.close()
    return filetype, numlines


def read_csv(inputFilePath, outFilePath, headerFields, fileEncoding, detect):
    encoding = "utf-8"
    outfile = io.open(outFilePath, 'w', encoding=encoding)

    if fileEncoding is not None:
        encoding = fileEncoding
    filetype = ""
    numlines = 0
    with io.open(inputFilePath, "rU", encoding=encoding) as file:
        sniffed = csv.Sniffer().sniff(file.readline())
        while sniffed.delimiter not in ";,\t":
            sniffed = csv.Sniffer().sniff(file.readline())
        file.seek(0)
        outData = ""
        reader = csv.reader(file, delimiter=sniffed.delimiter)
        if detect == "true":
            for i in reader:
                row = ';'.join(i) + '\n'
                if filetype == "":
                    filetype, header = getFileType(row, headerFields, encoding)
                    if header:
                        outData = header
                else:
                    break
        else:
            for i in reader:
                row = ';'.join(i) + '\n'
                if filetype == "":
                    filetype, header = getFileType(row, headerFields, encoding)
                    if header:
                        outData = header
                else:
                    outData = row
                    numlines = numlines + 1
            outfile.write(outData)
            outfile.close()
    return filetype, numlines


def filter_to_csv(inputFilePath, extension, outputFilePath, headerFields, fileEncoding, detect):
    filetype = ""
    numlines = 0

    try:
        if inputFilePath.lower().endswith(".csv") or extension == "csv":
            filetype, numlines = read_csv(inputFilePath, outputFilePath, headerFields, fileEncoding, detect)
        else:
            filetype, numlines = read_excel_xlrd(inputFilePath, outputFilePath, headerFields, detect)
    except Exception:
        filetype, numlines = read_csv(inputFilePath, outputFilePath, headerFields, fileEncoding, detect)

    return filetype, numlines

if __name__ == "__main__":
  parser = ArgumentParser(usage="%%prog -i infile -o outfile -hf headersFile")
  parser.add_argument('-i', "--infile", dest="inputFile", required=True, help="full path to file of the source spreadsheet")
  parser.add_argument('-d', "--detectMode", dest="detect", help="detect mode")
  parser.add_argument('-ex', "--extension", dest="extension", help="file extension")
  parser.add_argument('-en', "--fileEncoding", dest="fileEncoding", help="file encoding")
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
      if args.outputFile is None:
        args.outputFile = None
      else:
            if os.path.exists(args.outputFile) and not os.path.isfile(args.outputFile):
                basename = os.path.basename(args.inputFile)
                args.outputFile = args.outputFile + os.path.sep + os.path.splitext(basename)[0]
      if not os.path.isfile(args.inputFile):
        parser.error("File " + os.path.abspath(args.inputFile) + " is a folder")
      if not os.path.isfile(args.headersFile):
        parser.error("File " + os.path.abspath(args.headersFile) + " is a folder")
      if not args.outputFile or args.outputFile == "undefined":
        args.outputFile = args.inputFile + ".csv"

      fileEncoding = os.popen("file -b --mime-encoding " + args.inputFile).read()
      # print fileEncoding
      filetype, numLines = filter_to_csv(inputFilePath = args.inputFile, extension = args.extension, outputFilePath = args.outputFile, headerFields = args.headersFile, fileEncoding = fileEncoding, detect = args.detect)
      sys.stdout.write(json.JSONEncoder().encode({"fileType":filetype,"numLines":numLines,"outFile":os.path.abspath(args.outputFile)}))
