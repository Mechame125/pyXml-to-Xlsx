# SE NECESITA TENER ESTAS LIBRERIAS
import datetime
from html import parser
from os import path
import os
from xml.etree.ElementTree import PI, ParseError, parse
from xmlrpc.client import PARSE_ERROR
import xml.etree.ElementTree as ET
from openpyxl import Workbook
from os import path
import pathlib
from humanize import naturalsize

print()
print('  *******  EXTRAER DATOS DE .XML A .XLSX******* ')
print()

ruta = "D:\\Documentos\\Examples\\pyLeerXML\\"
fileXml = os.listdir(ruta)
print(' ARHIVOS  XML AGREGADOS ')
print(fileXml)
print()

results = []

wb = Workbook()   # Archivo de excel
ws = wb.active

for file in fileXml:
    # Llama todos los archivos que terminen en .xml
    if (file.endswith('.xml')):
      lstId = []
      p = pathlib.Path(file).absolute()     
      tree = ET.parse(p)
      sz = os.stat(p).st_size
      szk = sz/1024
      szkb = int(szk) # Peso en enteros KB del archivo
      root = tree.getroot()
      noPsn = 0   # Valor cantidad de personas
      dpMp = 0    # Valor id del lugar
      noSpv = 0   # Valor id de servidor
      noRcl = 0   # Valor id de cliente
      noHg = 0    # Valor cantidad de hogares 
      fch = ''    # Valor fecha actual
      
      # Busca los valores desde la parte externa hacia la parte interna
      for T_1 in root.findall('T_1'):
        dpMp = T_1.find('P2R1C1').text
        fch = 'INFORME '+datetime.date.today().strftime("%d-%m-%Y")

        for T_1_2 in root.findall('T_1_2'):
          noRcl = T_1_2.find('P7').text
      
          for T_1_3 in root.findall('T_1_3'):
            noSpv = T_1_3.find('P7S2').text
      
            for Hogares_count in root.findall('Hogares_count'):
              noHg = Hogares_count.find('P7S3').text

              for Hogares in root.findall('Hogares'):
                  noPsn = Hogares.find('Registro_count').text
                  
      # Muestro en pantalla los valores obtenidos
      print(dpMp, noSpv, noRcl, noHg, noPsn, file, szkb,fch)
      
      # En la hoja results organizo el orden como me quedarán guardados los valores
      results = [[dpMp, noSpv, noRcl, noHg, noPsn, file, szkb, fch]]
      
      # El resultado de un archivo queda guardado en una fila
      for row in results:
        ws.append(row)

      # Guardamos los valores en el excel
      wb.save('temporal_info.xlsx') 
      print()

print('Listo')
