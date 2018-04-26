import PyPDF2
import re
import openpyxl

pdfFileObj = open(r'C:\Users\Sun Shanbin\Desktop\PDF\test.pdf.','rb')

OPinfo=[]
OPinfoCon ={'Page':'Page','TESTDATA':'TESTDATA',}
OPinfo.append(OPinfoCon)



pdfReader =PyPDF2.PdfFileReader(pdfFileObj)
#print(pdfReader.numPages)

Page = 1
pageObj = pdfReader.getPage(Page)
pageCont=pageObj.extractText()

print(pageCont + '\n')

#TO DO Search Date
matchObj = re.search(r'Date:.*? ([0-9].*[0-9])\S', pageCont, re.M|re.I)
if matchObj:
    print('TEST data:', matchObj.group(1))
else:
    print('Date no match')

#TO DO Search File name
matchObj = re.search(r'Lab\n(.*?)DUT: ', pageCont, re.M|re.I)
if matchObj:
    print('File name:', matchObj.group(1))
else:
    print('File name no match')

#TO DO Search Medium fre:
matchObj = re.search(r'f = (.*?) MHz;', pageCont, re.M|re.I)
if matchObj:
    print('Medium usede f =:', matchObj.group(1)+' MHz')
else:
    print('Medium fre no match')

#TO DO Search Medium parameters:
matchObj = re.search(r' (.*?) S/m;.*?= (.*?);', pageCont, re.M|re.I)
if matchObj:
    print('Medium conductivy f =:', matchObj.group(1)+' S/m')
    print('Medium permittiy:', matchObj.group(2))
    # print('Medium permittiy:', matchObj.group(3))
else:
    print('Medium parameters no match')
#
# #TO DO Search Frequency:
# matchObj = re.search(r'Frequency: (.*?);Duty Cycle: ([0-9].*?)M', pageCont, re.M|re.I)
# if matchObj:
#     print('Frequency:', matchObj.group(1))
#     print('Duty Cycle:', matchObj.group(2))
# else:
#     print('Frequency Duty Cycle no match')
#
#
# #TO DO Search Probe
# matchObj = re.search(r'(EX3DV4|ES3DV3).*?SN([0-9]{4}).*?Calibrated.*?([0-9].*?);', pageCont, re.M|re.I)
# if matchObj:
#     print('Probe type:', matchObj.group(1))
#     print('Probe SN:', matchObj.group(2))
#     print('Probe Calibrate DATE:', matchObj.group(3))
# else:
#     print('Porbe no match')
#
# #TO DO Search Sensor-Surface
# matchObj = re.search(r'Sensor-Surface.*? (.*?mm)', pageCont, re.M|re.I)
# if matchObj:
#     print('Sensor-Surface:', matchObj.group(1))
# else:
#     print('Porbe no match')
#
# #TO DO Search DAE
# matchObj = re.search(r'DAE4.*?SN([0-9]+).*?Calibrated.*?([0-9].*?) ', pageCont, re.M|re.I)
# if matchObj:
#     print('DAE SN:', matchObj.group(1))
#     print('DAE Calibrate DATE:', matchObj.group(2))
# else:
#     print('DAE no match')
#
# #TO DO Search Phantom
# matchObj = re.search(r'Phantom.*? (.*)?;.*?Type.*? (.*?);.*Serial.*? (.*?[0-9]{4})', pageCont, re.M|re.I)
# if matchObj:
#     print('Phantom:', matchObj.group(1))
#     print('Phantom Type:', matchObj.group(2))
#     print('Phantom Serial:', matchObj.group(3))
# else:
#     print('Phantom no match')
#
# #TO DO Search Area Scan dx
# matchObj = re.search(r'Area Scan.*?dx=([0-9]+?mm).*?dy=([0-9]+?mm)', pageCont, re.M|re.I)
# if matchObj:
#     print('Area Scan dx=:', matchObj.group(1))
#     print('Area Scan dy=:', matchObj.group(2))
# else:
#     print('Area Scan dx no match')
#
# #TO DO Search Zoom Scan dx
# matchObj = re.search(r'Zoom Scan.*?dx=([0-9]{1,2}mm).*?dy=([0-9]{1,2}mm).*?dz=([0-9]{1,2}mm)', pageCont, re.M|re.I)
# if matchObj:
#     print('Zoom Scan dx=:', matchObj.group(1))
#     print('Zoom Scan dy=:', matchObj.group(2))
#     print('Zoom Scan dz=:', matchObj.group(3))
# else:
#     print('Zoom Scan dx no match')
#
#
#
# #TO DO Search SAR and drift
# matchObj = re.search(r'[(]1 g[)] = (.*?) .*?[(]10 g[)] = (.*?) ' , pageCont, re.M|re.I)
# if matchObj:
#     print('1-g SAR=:', matchObj.group(1))
#     print('10-g SAR=:', matchObj.group(2))
# else:
#     print('SAR value  no match')
#
# #TO DO Search SAR and drift
# matchObj = re.search(r'Power Drift = (.*?) ' , pageCont, re.M|re.I)
# if matchObj:
#     print('Power drift dx=:', matchObj.group(1))
# else:
#     print('Power drift no match')

