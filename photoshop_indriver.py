import win32com.client
import os
import pandas as pd
from pandas import DataFrame

asistencia1 = "C:/Users/ASESOR 5B/OneDrive/Escritorio/16enero.xlsx"
a1 = pd.read_excel(asistencia1, sheet_name='asistencia')
l1 = (a1["NOMBRE"].tolist())

asistencia2 = "C:/Users/ASESOR 5B/OneDrive/Escritorio/15enero.xlsx"
a2 = pd.read_excel(asistencia2, sheet_name='asistencia')
l2 = (a2["NOMBRE"].tolist())

def nombres_en_comun (l1, l2):
    conjunto1 = set(l1)
    conjunto2 = set(l2)

    return conjunto1 & conjunto2

nombres = nombres_en_comun(l1, l2)

psApp = win32com.client.Dispatch("Photoshop.Application")
psApp.Open(r"C:/Users/ASESOR 5B/OneDrive/Escritorio/reconocimiento.psd")
doc = psApp.Application.ActiveDocument

for text in nombres:
    layerText = doc.ArtLayers["TextoEditable"]
    text_of_layer = layerText.TextItem
    text_of_layer.contents = text

    options = win32com.client.Dispatch("Photoshop.ExportOptionsSaveForWeb")
    options.Format = 13
    options.PNG8 = False

    pngfile = f"C:/Users/ASESOR 5B/OneDrive/Escritorio/export/{text}.png"
    doc.Export(ExportIn=pngfile, ExportAs=2, Options = options)
