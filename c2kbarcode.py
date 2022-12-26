import xlwt
from datetime import datetime

def main():
    skiplines = 2                     #the number of lines at the start of the line to skip
    filename = 'Pallet #3B-SDUSD.txt' #the name of the file to process
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
    if (line.startswith('NO SERIAL NUMBER 11 E')):
        return 'Yoga 11e'
    if (line.startswith('C00') or line.startswith('NO ASSET TAG') or line.startswith('NO SERIAL')):
        return 'n/a'
    if (line.startswith('http://s')):
        return 'N22-20'
    if (line.startswith('MTM') or line.startswith('LR0')):
        return 'N23'
    if (line.startswith('1S20')):
        return 'Yoga 11e'
    raise Exception('model: ' + line + 'is an unknown format')

def ParseLineSN(line):
    sampleLRline = 'LR0AURANLRNXB822600D'
    sampleCline = 'C000528173'
    notagline = 'NO ASSET TAG'
    line = line.strip()
    if (line.startswith('LR') and (len(line) == len(sampleLRline))):  #easy - serial like LRxxxx
        return line
    if (line.startswith('1S2') and (len(line) == len(sampleLRline))):  #yoga 11E
        return line
    if ((line.startswith('C00') and len(line) == len(sampleCline)) or line == notagline):   #easy - asset tag starting with C
        return line
    if (line.startswith('NO SERIAL NUMBER')):
        return line
    if ('S/N:' in line and 'MO:' in line and 'MTM:' in line):
        snpos = line.index('S/N:') + 3
        mopos = line.index('MO:') + 3
        moend = line.index('MTM:')
        return (line[snpos:mopos-3].strip() + line[mopos:moend].strip()).replace(':','').replace(';','').replace(',','')
    if ('SN,' in line and 'MO,' in line and 'MTM,' in line):
        snpos = line.index('SN,') + 3
        mopos = line.index('MO,') + 3
        return (line[snpos:mopos-3].strip() + line[mopos:].strip()).replace(':','').replace(';','').replace(',','')
    raise Exception(line + 'is an unknown format')

main()