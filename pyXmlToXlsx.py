'''
from html import parser
from os import path
import os
from xml.etree.ElementTree import PI, ParseError, parse
from xmlrpc.client import PARSE_ERROR
#from lxml import etree
#from openpyxl import Workbook
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from os import path
#import pandas as pd
import pathlib
from humanize import naturalsize

print()
print('  *******  EXTRAER DATOS DE .XML ******* ')
print()

ruta = "D:\\Documentos\\MECH\\DANE\\pyLeerXML\\"

fileXml = os.listdir(ruta)
print(' ARHIVOS  XML AGREGADOS ')
print(fileXml)
print()

results = []

wb = Workbook()
ws = wb.active
#res = []

for file in fileXml:
    if (file.endswith('.xml')):
      print('file ', file)
      #res = [file]
      pl = os.path.abspath(file)
      p = pathlib.Path(file).absolute()
      noPsn = 0
      rslt = 'EC - ENCUESTA COMPLETA'
      #print('p.name ', p.name)
      #print('p.parent ', p.parent)
      #print('p ', p)
      tree = ET.parse(p)
      sz = os.stat(p).st_size
      szk = sz/1024
      szkb = int(szk)
      #naturalsize(sz)
      root = tree.getroot()
      dpMp = 0
      noSgm = 0
      noMzn = 0
      noSpv = 0
      noRcl = 0
      noEdf = 0
      noViv = 0
      noHg = '1'
      fch = ''

      for T_1 in root.findall('T_1'):
        dpMp = T_1.find('P2R1C1').text
        noSgm = T_1.find('P7R1C2').text
        noMzn = T_1.find('P5R1C1').text
        noEdf = T_1.find('P1370').text
        noViv = T_1.find('P6').text
        fch = 'INFORME 19-09-2024'
        brr = T_1.find('I8R1C1').text
    
        for T_1_2 in root.findall('T_1_2'):
          noRcl = T_1_2.find('P7').text
      
          for T_1_3 in root.findall('T_1_3'):
            noSpv = T_1_3.find('P7S2').text
      
            for Hogares_count in root.findall('Hogares_count'):
              noHg = '1'

              for Hogares in root.findall('Hogares'):
                noPsn = Hogares.find('Registro_count').text
                rslt = 'EC - ENCUESTA COMPLETA'
            
      print(dpMp, noSpv, noRcl, noSgm, noMzn, noEdf, noViv, noHg, noPsn, rslt, file, szkb,fch)
      results = [[dpMp, noSpv, noRcl, noSgm, noMzn, noEdf, noViv, noHg, noPsn, rslt, file, szkb, fch]]
      print()
      for row in results:
        ws.append(row)

      wb.save('temporal_Encuesta.xlsx') 

print()
'''

'''
for Personas in root.findall('Personas'):
              contPsn = Personas.find('contador').text
              print('contador psn ', contPsn)
              for T_5 in root.findall('T_5'):                    
                noId = T_5.find('III3R2C3').text
                print(noId)       
                while noPsn == contPsn:
                    noId = T_5.find('III3R2C3').text
                else:
                    noIds = T_5.find('III3R2C3').text          
                
    contPsn = 1
                    #for Personas in root.findall('Personas'):
                contPsn = Hogares.find('contador').text
                if contPsn == 1:
                    for T_5 in root.findall('Personas'):
                        noId = T_5.find('III3R2C3').text
    
'''


from html import parser
from os import path
import os
from xml.etree.ElementTree import PI, ParseError, parse
from xmlrpc.client import PARSE_ERROR
#from lxml import etree
#from openpyxl import Workbook
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from os import path
#import pandas as pd
import pathlib
from humanize import naturalsize

print()
print('  *******  EXTRAER DATOS DE .XML ******* ')
print()

ruta = "D:\\Documentos\\MECH\\DANE\\pyLeerXML\\"
fileXml = os.listdir(ruta)
print(' ARHIVOS  XML AGREGADOS ')
print(fileXml)
print()

results = []
noIds = ()

wb = Workbook()
ws = wb.active

for file in fileXml:
    if (file.endswith('.xml')):
      print('file ', file)
      lstId = []
      contPsn = 1
      pl = os.path.abspath(file)
      p = pathlib.Path(file).absolute()
      noPsn = 0
      rslt = 'EC - ENCUESTA COMPLETA'
      #print('p.name ', p.name)
      #print('p.parent ', p.parent)
      #print('p ', p)
      tree = ET.parse(p)
      sz = os.stat(p).st_size
      szk = sz/1024
      szkb = int(szk)
      #naturalsize(sz)
      root = tree.getroot()
      dpMp = 0
      noSgm = 0
      noMzn = 0
      noSpv = 0
      noRcl = 0
      noEdf = 0
      noTel = 0
      noViv = 0
      noHg = '1'
      fch = ''
      lstPsn = ''
      noId = 0
      contPsn = 0

      for T_1 in root.findall('T_1'):
        dpMp = T_1.find('P2R1C1').text
        noSgm = T_1.find('P7R1C2').text
        noMzn = T_1.find('P5R1C1').text
        noEdf = T_1.find('P1370').text
        noViv = T_1.find('P6').text
        fch = 'INFORME 15-10-2024'
        brr = T_1.find('I8R1C1').text
        noTel = T_1.find('I10R1C1').text

        for T_1_2 in root.findall('T_1_2'):
          noRcl = T_1_2.find('P7').text
      
          for T_1_3 in root.findall('T_1_3'):
            noSpv = T_1_3.find('P7S2').text
      
            for Hogares_count in root.findall('Hogares_count'):
              noHg = '1'

              for Hogares in root.findall('Hogares'):
                  noPsn = Hogares.find('Registro_count').text
                  rslt = 'EC - ENCUESTA COMPLETA'
                  lstPsn = Hogares.find('lista1').text

      print(dpMp, noSpv, noRcl, noSgm, noMzn, noEdf, noViv, noHg, noPsn, rslt, file, szkb,fch, lstPsn, noTel)
      results = [[dpMp, noSpv, noRcl, noSgm, noMzn, noEdf, noViv, noHg, noPsn, rslt, file, szkb, fch, lstPsn, noTel]]
      
      for row in results:
        ws.append(row)

      wb.save('temporal_Encuesta.xlsx') 
      print()

print('Listo')
