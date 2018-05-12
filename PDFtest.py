import PyPDF2
import re
import openpyxl
from openpyxl.workbook import Workbook
from openpyxl.writer.excel import ExcelWriter
from openpyxl.cell import Cell

EquipmentList=[]
EquipmentListCon ={'Manufacturer':'SPEAG', 'Device':'Device','Type':'Type','SN':'Serial number','DateOfCal':'Date of last calibration','Vaild period':'Valid period'}
EquipmentList.append(list(EquipmentListCon))
EquipmentListConDAE =[]
EquipmentListConPROBE =[]
EquipmentListConPHANTOM =[]


PageConList=[]
PageCon ={'TEST data':'TEST data', 'File name':'File name','DUT':'DUT','Type':'Type', 'Serial':'Serial',\
          'Medium used f':'Medium used f','Medium conductivy':'Medium conductivy','Medium permittiy':'Medium permittiy',\
          'Frequency':'Frequency','Duty Cycle':'Duty Cycle','Probe type':'Probe type','Probe SN':'Probe SN','Probe Calibrate DATE':'Probe Calibrate DATE',\
          'Sensor-Surface':'Sensor-Surface','DAE SN':'DAE SN','DAE Calibrate DATE':'DAE Calibrate DATE',\
           'Phantom':'Phantom','Phantom Type':'Phantom Type','Phantom Serial':'Phantom Serial',\
           'Area Scan dx':'Area Scan dx','Area Scan dy':'Area Scan dy','Zoom Scan dx':'Zoom Scan dx','Zoom Scan dy':'Zoom Scan dy','Zoom Scan dz':'Zoom Scan dz',\
           '1-g SAR':'1-g SAR','10-g SAR':'10-g SAR','Power drift':'Power drift'}
PageConList.append(list(PageCon.values()))

pdfFileObj = open(r'C:\Users\Sun Shanbin\Desktop\PDF\test.pdf.', 'rb')
pdfReader =PyPDF2.PdfFileReader(pdfFileObj)
#print(pdfReader.numPages)


# print(pageCont + '\n')

def TiQuInfo(Page):
    pageObj = pdfReader.getPage(Page)
    pageCont = pageObj.extractText()

    # if Page == 4:
    #     print(pageCont)

    #TO DO Search Date
    matchObj = re.search(r'Date:.*? ([0-9].*[0-9])\S', pageCont, re.M|re.I)
    if matchObj:
        # print('TEST data:', matchObj.group(1))
        PageCon['TEST data'] = matchObj.group(1)
    else:
        # print('Date no match')
        PageCon['TEST data'] = 'Date no match'


    #TO DO Search File name
    matchObj = re.search(r' Lab\n(.*?)DUT: ', pageCont, re.M|re.S|re.I)
    if matchObj:
        print('File name:', matchObj.group(1))
        PageCon['File name'] = matchObj.group(1)
    else:
        print('File name no match')
        PageCon['File name'] ='File name no match'


    #TO DO Search DUT
    matchObj = re.search(r'DUT: (.*?); Type: (.*?); Serial: (.*?)\s', pageCont, re.M|re.I)
    if matchObj:
        PageCon['DUT'] = matchObj.group(1)
        PageCon['Type'] = matchObj.group(2)
        PageCon['Serial'] = matchObj.group(3)
    else:
        PageCon['DUT'] = 'DUT name no match'
        PageCon['Type'] = 'DUT name no match'
        PageCon['Serial'] = 'DUT name no match'

    #TO DO Search Medium fre:
    matchObj = re.search(r'f = (.*?) MHz;', pageCont, re.M|re.I)
    if matchObj:
        print('Medium used f =', matchObj.group(1)+' MHz')
        PageCon['Medium used f'] = matchObj.group(1)
    else:
        print('Medium fre no match')
        PageCon['Medium used f'] = 'Medium fre no match'


    #TO DO Search Medium parameters:
    matchObj = re.search(r'= (\d.\d*?) S/m;.*?= (.*?);', pageCont, re.M|re.I)
    if matchObj:
        print('Medium conductivy f =', matchObj.group(1)+' S/m')
        print('Medium permittiy:', matchObj.group(2))
        # print('Medium permittiy:', matchObj.group(3))
        PageCon['Medium conductivy'] = matchObj.group(1)
        PageCon['Medium permittiy'] = matchObj.group(2)
    else:
        print('Medium parameters no match')
        PageCon['Medium conductivy'] = 'Medium parameters no match'
        PageCon['Medium permittiy'] = 'Medium parameters no match'


    #TO DO Search Frequency:
    matchObj = re.search(r'Frequency: (.*?);Duty Cycle: ([0-9].*?)M', pageCont, re.M|re.S|re.I)
    if matchObj:
        print('Frequency:', matchObj.group(1))
        print('Duty Cycle:', matchObj.group(2))
        PageCon['Frequency'] = matchObj.group(1)
        PageCon['Duty Cycle'] = matchObj.group(2)
    else:
        print('Frequency Duty Cycle no match')
        PageCon['Frequency'] = 'Frequency & Duty Cycle no match'
        PageCon['Duty Cycle'] = 'Frequency & Duty Cycle no match'


    #TO DO Search Probe
    matchObj = re.search(r'(EX3DV4|ES3DV3).*?SN([0-9]{4}).*?Calibrated: ([0-9].*?);',pageCont, re.M|re.S|re.I)
    if matchObj:
        print('Probe type:', matchObj.group(1))
        print('Probe SN:', matchObj.group(2))
        print('Probe Calibrate DATE:', matchObj.group(3))
        PageCon['Probe type'] = matchObj.group(1)
        PageCon['Probe SN'] = matchObj.group(2)
        PageCon['Probe Calibrate DATE'] = matchObj.group(3)

        if  matchObj.group(2) not in EquipmentListConPROBE:
            EquipmentListConPROBE.append(matchObj.group(2))
            EquipmentListCon = {'Manufacturer': 'SPEAG', 'Device': 'Dosimetric E-Field Probe', 'Type': PageCon['Probe type'], 'SN': PageCon['Probe SN'],
                                'DateOfCal': PageCon['Probe Calibrate DATE'], 'Vaild period': 'One year'}
            EquipmentList.append(list(EquipmentListCon.values()))

    else:
        print('Porbe no match')
        PageCon['Probe type'] = 'Porbe no match'
        PageCon['Probe SN'] = 'Porbe no match'
        PageCon['Probe Calibrate DATE'] = 'Porbe no match'


    #TO DO Search Sensor-Surface
    matchObj = re.search(r'Sensor-Surface.*? (.*?mm)', pageCont, re.M|re.I)
    if matchObj:
        print('Sensor-Surface:', matchObj.group(1))
        PageCon['Sensor-Surface'] = matchObj.group(1)
    else:
        print('Sensor-Surface: no match')
        PageCon['Sensor-Surface'] = 'Sensor-Surface: no match'


    #TO DO Search DAE
    matchObj = re.search(r'DAE4.*?SN([0-9]+).*?Calibrated.*?([0-9].*?) ', pageCont, re.M|re.I)
    if matchObj:
        print('DAE SN:', matchObj.group(1))
        print('DAE Calibrate DATE:', matchObj.group(2))
        PageCon['DAE SN'] = matchObj.group(1)
        PageCon['DAE Calibrate DATE'] = matchObj.group(2)
        if  matchObj.group(2) not in EquipmentListConDAE:
            EquipmentListConDAE.append(matchObj.group(2))

            EquipmentListCon = {'Manufacturer': 'SPEAG', 'Device': 'Data acquisition electronics', 'Type': 'DAE4', 'SN': PageCon['DAE SN'],
                                'DateOfCal': PageCon['DAE Calibrate DATE'], 'Vaild period': 'One year'}
            EquipmentList.append(list(EquipmentListCon.values()))
    else:
        print('DAE no match')
        PageCon['DAE SN'] = 'DAE no match'
        PageCon['DAE Calibrate DATE'] = 'DAE no match'


    #TO DO Search Phantom
    matchObj = re.search(r'Phantom.*? (.*)?;.*?Type.*? (.*?);.*Serial.*? (.*?[0-9]{4})', pageCont, re.M|re.I)
    if matchObj:
        print('Phantom:', matchObj.group(1))
        print('Phantom Type:', matchObj.group(2))
        print('Phantom Serial:', matchObj.group(3))
        PageCon['Phantom'] = matchObj.group(1)
        PageCon['Phantom Type'] = matchObj.group(2)
        PageCon['Phantom Serial'] = matchObj.group(3)
        if  matchObj.group(3) not in EquipmentListConPHANTOM:
            EquipmentListConPHANTOM.append(matchObj.group(3))
            EquipmentListCon = {'Manufacturer': 'SPEAG', 'Device': 'Twin Phantom ', 'Type': PageCon['Phantom Type'], 'SN': PageCon['Phantom Serial'],
                                'DateOfCal': 'NCR', 'Vaild period': 'NCR'}
            EquipmentList.append(list(EquipmentListCon.values()))
    else:
        print('Phantom no match')
        PageCon['Phantom'] = 'Phantom no match'
        PageCon['Phantom Type'] = 'Phantom no match'
        PageCon['Phantom Serial'] = 'Phantom no match'


    #TO DO Search Area Scan dx
    matchObj = re.search(r'Area Scan.*?dx=([0-9]+?mm).*?dy=([0-9]+?mm)', pageCont, re.M|re.I)
    if matchObj:
        print('Area Scan dx=:', matchObj.group(1))
        print('Area Scan dy=:', matchObj.group(2))
        PageCon['Area Scan dx'] = matchObj.group(1)
        PageCon['Area Scan dy'] = matchObj.group(2)
    else:
        print('Area Scan dx no match')
        PageCon['Area Scan dx'] = 'Area Scan dx no match'
        PageCon['Area Scan dy'] = 'Area Scan dx no match'

    #TO DO Search Zoom Scan dx
    matchObj = re.search(r'Zoom Scan.*?dx=([0-9]{1,2}mm).*?dy=([0-9]{1,2}mm).*?dz=([0-9]{1,2}mm)', pageCont, re.M|re.I)
    if matchObj:
        print('Zoom Scan dx=:', matchObj.group(1))
        print('Zoom Scan dy=:', matchObj.group(2))
        print('Zoom Scan dz=:', matchObj.group(3))
        PageCon['Zoom Scan dx'] = matchObj.group(1)
        PageCon['Zoom Scan dy'] = matchObj.group(2)
        PageCon['Zoom Scan dz'] = matchObj.group(3)
    else:
        print('Zoom Scan dx no match')
        PageCon['Zoom Scan dx'] = 'Zoom Scan dx no match'
        PageCon['Zoom Scan dy'] = 'Zoom Scan dx no match'
        PageCon['Zoom Scan dz'] = 'Zoom Scan dx no match'


    #TO DO Search SAR and drift
    matchObj = re.search(r'[(]1 g[)] = (.*?) .*?[(]10 g[)] = (.*?) ' , pageCont, re.M|re.I)
    if matchObj:
        print('1-g SAR=:', matchObj.group(1))
        print('10-g SAR=:', matchObj.group(2))
        PageCon['1-g SAR'] = matchObj.group(1)
        PageCon['10-g SAR'] = matchObj.group(2)
    else:
        print('SAR value  no match')
        PageCon['1-g SAR'] = 'SAR value  no match'
        PageCon['10-g SAR'] = 'SAR value  no match'


    #TO DO Search SAR and drift
    matchObj = re.search(r'Power Drift = (.*?) ' , pageCont, re.M|re.I)
    if matchObj:
        print('Power drift =:', matchObj.group(1))
        PageCon['Power drift'] = matchObj.group(1)
    else:
        print('Power drift no match')
        PageCon['Power drift'] = 'Power drift no match'

    PageConList.append(list(PageCon.values()))

    # print('\n', PageConList)

    # print('\n', EquipmentList)

def prtexcel():
    wb = Workbook()
    ws = wb['Sheet']

    for b in range(0 ,len(PageConList)):
        ws.append(PageConList[b])

    ws1 = wb.create_sheet()

    for b in range(0 ,len(EquipmentList)):
        ws1.append(EquipmentList[b])

    wb.save(r'C:\Users\Sun Shanbin\Desktop\PDF\TEST.xlsx')


# for Page in range(0,pdfReader.numPages):
for Page in range(0,pdfReader.numPages):
    TiQuInfo(Page)

prtexcel()