import xlwt
from datetime import datetime

def main():
    skiplines = 1                                   #the number of lines at the start of the line to skip
    filename = 'c:\c2kbarcode\SDUSD (12-15-22).txt' #the name of the file to process
    textfile = open(filename,'r')
    # fix up the input lines based on the different barcode formats
    lines = []
    for line in textfile:
        if (skiplines > 0):  # skip the number of lines that are in the header
                skiplines-= 1
                continue
        if not line.strip(): # detect empty line
            continue
        # print(ParseLine(line)) for debugging
        lines.append(ParseLine(line))
    # add the fixed up stuff to a new Excel document
    excelFilename = filename.replace(".txt",".xls")
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Scanned Items')
    row = 0
    i = 0
    while i < len(lines)-1:
        ws.write(row, 0,lines[i])
        i = i + 1
        ws.write(row, 1,lines[i])
        row = row +1
        i = i+1
    wb.save(excelFilename)


def ParseLine(line):
    sampleLRline = 'LR0AURANLRNXB822600D'
    sampleCline = 'C000528173'
    notagline = 'NO ASSET TAG'
    line = line.strip()
    if (line.startswith('LR') and (len(line) == len(sampleLRline))):  #easy - serial like LRxxxx
        return line
    if (line.startswith('1S2') and (len(line) == len(sampleLRline))):  #easy - serial like 1S2xxx, same length as above
        return line
    if ((line.startswith('C00') and len(line) == len(sampleCline)) or line == notagline):   #easy - asset tag starting with C
        return line
    if (line.startswith('NO SERIAL NUMBER')):
        return line
    if ('S/N:' in line and 'MO:' in line and 'MTM:' in line):
        mopos = line.index('MO:') + 3
        moend = line.index('MTM:')
        return line[mopos:moend].strip()
    if ('SN,' in line and 'MO,' in line and 'MTM,' in line):
        mopos = line.index('MO,') + 3
        return line[mopos:].strip()
    raise Exception(line + 'is an unknown format')

main()

        

# style0 = xlwt.easyxf('font: name Times New Roman, color-index red, bold on',
#     num_format_str='#,##0.00')
# style1 = xlwt.easyxf(num_format_str='D-MMM-YY')

# wb = xlwt.Workbook()
# ws = wb.add_sheet('A Test Sheet')

# ws.write(0, 0, 1234.56, style0)
# ws.write(1, 0, datetime.now(), style1)
# ws.write(2, 0, 1)
# ws.write(2, 1, 1)
# ws.write(2, 2, xlwt.Formula("A3+B3"))

# wb.save('c:\c2kbarcode\example2.xls')
