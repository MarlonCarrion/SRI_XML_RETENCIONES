import os
import xml.etree.ElementTree as ET
import pandas as pd
import xlsxwriter

path = 'D:/Escritorio/Retenciones Devoluciones Prueba/XML/'
writer = pd.ExcelWriter(path+'Retenciones.xlsx', engine='xlsxwriter')
row = {}
df = pd.DataFrame(row)
df_details = pd.DataFrame(row)

dirs = os.listdir(path)
for file in dirs:
    archivo, ext = os.path.splitext(file)
    if ext != '.xml':
        print('File: ' + str(file) + str(' Is not XML'))
    else:
        tree = ET.parse(path+file)
        root = tree.getroot()
        for child in root:
            if child.tag =='comprobante':
                comprobante = child.text
                c = ET.fromstring(comprobante)
                h, k = {}, {}
                for x in range(len(c)):
                    if c[x].tag == 'infoTributaria':
                        for i in c[x]:
                            h[i.tag] = i.text
                    if c[x].tag == 'infoCompRetencion':
                        for i in c[x]:
                            h[i.tag] = i.text
                    if c[x].tag == 'impuestos':
                        for i in c[x]:
                            for j in i:
                                if j.tag == 'numDocSustento':
                                    h[j.tag] = j.text
                                k[j.tag] = j.text
                            df_details = df_details.append(k, ignore_index=True)

                df = df.append(h, ignore_index=True)
                df = df[['fechaEmision','ruc','razonSocial','claveAcceso','estab','ptoEmi','secuencial','numDocSustento']]
                df_details = df_details[['numDocSustento','codigo','codigoRetencion','baseImponible','porcentajeRetener',
                                         'valorRetenido']]
#Convertir en numero
df_details['baseImponible'] = pd.to_numeric(df_details['baseImponible'])
df_details['porcentajeRetener'] = pd.to_numeric(df_details['porcentajeRetener'])
df_details['valorRetenido'] = pd.to_numeric(df_details['valorRetenido'])
#Convertir en fecha
df['fechaEmision'] = pd.to_datetime(df['fechaEmision'], format='%d/%m/%Y').dt.date
#Guardar en hoja de excel
df.to_excel(writer, index=False, sheet_name='General')
df_details.to_excel(writer, index=False, sheet_name='Detalle')
#Bordes#
workbook = writer.book
worksheet = writer.sheets['General']
worksheet_details = writer.sheets['Detalle']
border_fmt = workbook.add_format({'bottom':1, 'top':1, 'left':1, 'right':1})
worksheet.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df), len(df.columns)-1), {'type': 'no_errors',
                                                                                             'format': border_fmt})
worksheet_details.conditional_format(xlsxwriter.utility.xl_range(0, 0, len(df_details), len(df_details.columns)-1),
                                     {'type': 'no_errors', 'format': border_fmt})
print("******************************\n Archivo Generado Exitosamente \n******************************")
writer.save()