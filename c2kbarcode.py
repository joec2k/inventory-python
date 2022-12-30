import xlwt
from datetime import datetime

def main():
    skiplines = 1                     #the number of lines at the start of the line to skip
    filename = 'SDUSD 2-of-2 (12-15-22).txt' #the name of the file to process
    textfile = open(filename,'r')
    # fix up the input lines based on the different barcode formats
    lines = []
    modellines = []
    for line in textfile:
        if (skiplines > 0):  # skip the number of lines that are in the header
                skiplines-= 1
                continue
        if not line.strip(): # detect empty line
            continue
        line = line.lower()
        # print(ParseLine(line)) for debugging
        lines.append(ParseLineSN(line))
        modellines.append(ParseLineModel(line))
    # add the fixed up stuff to a new Excel document
    excelFilename = filename.replace(".txt",".xls")
    wb = xlwt.Workbook()
    ws = wb.add_sheet('Scanned Items')
    #write a header
    ws.write(0, 0, 'Asset Type (Model)')
    ws.write(0, 1, 'Serial Number')
    ws.write(0, 2, 'Asset Tag')
    #write the body
    row = 1
    i = 0
    while i < len(lines)-1:
        ws.write(row, 0, modellines[i])
        ws.write(row, 1,lines[i])
        i = i + 1
        ws.write(row, 2,lines[i])
        row = row +1
        i = i+1
    wb.save(excelFilename)


def ParseLineModel(line):
    if (line.startswith('NO SERIAL NUMBER 11 E'.lower())):
        return 'Yoga 11e'
    if (line.startswith('C00'.lower()) or line.startswith('NO ASSET TAG'.lower()) or line.startswith('NO SERIAL'.lower())):
        return 'n/a'
    if (line.startswith('http://s'.lower())):
        return 'N22-20'
    if (line.startswith('MTM'.lower()) or line.startswith('LR0'.lower())):
        return 'N23'
    if (line.startswith('1S20'.lower())):
        return 'Yoga 11e'
    raise Exception('model: ' + line + 'is an unknown format')

def ParseLineSN(line):
    sampleLRline = 'LR0AURANLRNXB822600D'.lower()
    sampleCline = 'C000528173'.lower()
    notagline = 'NO ASSET TAG'.lower()
    line = line.strip()
    if (line.startswith('LR'.lower()) and (len(line) == len(sampleLRline))):  #easy - serial like LRxxxx
        return line
    if (line.startswith('1S2'.lower()) and (len(line) == len(sampleLRline))):  #yoga 11E
        return line
    if ((line.startswith('C00'.lower()) and len(line) == len(sampleCline)) or line == notagline):   #easy - asset tag starting with C
        return line
    if (line.startswith('NO SERIAL NUMBER'.lower())):
        return line
    if ('S/N:'.lower() in line and 'MO:'.lower() in line and 'MTM:'.lower() in line):
        snpos = line.index('S/N:'.lower()) + 3
        mopos = line.index('MO:'.lower()) + 3
        moend = line.index('MTM:'.lower())
        return (line[snpos:mopos-3].strip() + line[mopos:moend].strip()).replace(':','').replace(';','').replace(',','')
    if ('SN,'.lower() in line and 'MO,'.lower() in line and 'MTM,'.lower() in line):
        snpos = line.index('SN,'.lower()) + 3
        mopos = line.index('MO,'.lower()) + 3
        return (line[snpos:mopos-3].strip() + line[mopos:].strip()).replace(':','').replace(';','').replace(',','')
    raise Exception(line + ' is an unknown format')

main()