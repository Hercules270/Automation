from camelot import read_pdf
from ocrmypdf import ocr
import sys, getopt


short_options = 'i:o:'
long_options = ['input=', 'output=']

try:
    arguments, values = getopt.getopt(sys.argv[1:], short_options, long_options)
except getopt.error as err:
    # Output error, and return with an error code
    print (str(err))
    sys.exit(2)

for current_argument, current_value in arguments:
    if current_argument in ("-i", "--input"):
        inputFile = current_value
    elif current_argument in ("-o", "--output"):
        outputFile = '/mnt/d/GitHub/Automation/Pdf Converter/' + current_value

ocr(inputFile, outputFile, deskew = True)
tables = read_pdf(outputFile, edge_tol = 500, flavor = 'stream')
tables.export(outputFile[:-3] + 'xlsx', f = 'excel')




