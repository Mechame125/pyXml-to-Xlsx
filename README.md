# pyXml-to-Xlsx
Este código funciona para leer y exportar información especifica de varios archivos .xml y pasarlos a un archivo .xlsx

Desde cualquier carpeta, el código primero lee la cantidad de documentos en la carpeta, 
    
                
            ruta = "D:\\Documentos\\Examples\\pyLeerXML\\"
            fileXml = os.listdir(ruta)

Después, en un archivo de excel existente para almacenar la información,
    
                
            wb = Workbook()  
            ws = wb.active

Se selecciona los archivos que tienen formato .xml y los lista,
    
                
            for file in fileXml:
                if (file.endswith('.xml')):
                    lstId = []

Luego, empieza a buscar la información que pedimos, 
    
                
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

Almacena la información en una lista, según el orden que nos guste,
    
                
              results = [[dpMp, noSpv, noRcl, noHg, noPsn, file, szkb, fch]]

Finalmente, guarda y exporta la información obtenida de todos los archivos .xml en el archivo .xlsx,
    
                
              wb.save('temporal_info.xlsx')
      
---
Mech.
