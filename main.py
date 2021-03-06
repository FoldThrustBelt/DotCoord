print("Ingrese el nombre del archivo KML a leer (ex. oLa.kml): ")
kmlinput = input()

# Importando módulos
from xml.dom import minidom
import xlsxwriter
import pandas as pd
import os

# Defino el spreadsheet
spreadsheet = xlsxwriter.Workbook("ExcelPuntos.xlsx")

# Defino el sheet del spreadsheet
sheet = spreadsheet.add_worksheet()

# Definiendo las columnas del spreadsheet
sheet.write("A1", "Object-ID")
sheet.write("B1", "Coordenada X")
sheet.write("C1", "Coordenada Y")


# Creo los array que voy a poblar
oid = []
foroid = 0
truex = []
truey = []


# Acá va el archivo que KML que se pretende leer
kml = minidom.parse(kmlinput)

# Usando una función del minidom
Placemarks = kml.getElementsByTagName("Placemark")

# Cuantos puntos hay en total?
coordquant = len(Placemarks)

# Transformé la string resultante en un dato tipo float (es necesario esto?)
coordquant = float(coordquant)

# Declaro un valor inicial para el loop (valor loopificador)
nmark = 0

# Mientras que mi valor sea menor al valor del total de puntos, realiza:
while nmark < coordquant:

# Usando el minidom para parsear las coordenadas
    coords = Placemarks[nmark].getElementsByTagName("coordinates")

# Dejando claro que solo me interesa ese nodo
    puntos = coords[0].firstChild.data

# Creando variables spliteando el resultado
    x, y, z = puntos.split(",")

# xyz ahora son datos tipo float, pero z no me interesa por ahora
    x = float(x)
    y = float(y)

# Metiendo los datos x,y de c/ punto en los array
    oid.append(foroid)
    truex.append(x)
    truey.append(y)
    sheet.write_column("A2", oid)
    sheet.write_column("B2", truex)
    sheet.write_column("C2", truey)
    print("Leyendo coordenadas: " + str(x) + "," + str(y))

# Valor loopificador gana 1
    nmark = nmark + 1
    foroid = foroid + 1

# Si no se cumple la condición del while, entonces terminó de leer
else:
    coordquant = str(coordquant)
    print("Terminé. Leí " + coordquant + " puntos.")
    spreadsheet.close()
    print("Generando archivo .csv.")
    read_file = pd.read_excel(r'ExcelPuntos.xlsx')
    read_file.to_csv(r'ExcelPuntosCSV.csv', index=None, header=True)
    os.remove("ExcelPuntos.xlsx")
    print("Archivos extra eliminados.")
