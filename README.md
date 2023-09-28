# Photoshop Driver Readme

## Descripción
Este script es una herramienta de automatización para Photoshop diseñada para agilizar el proceso de actualización de capas de texto y la exportación de imágenes basadas en datos de archivos de Excel. Es particularmente útil en escenarios donde tienes diferentes archivos de Excel que contienen nombres y deseas generar imágenes individualizadas en Photoshop.

## Características
1. **Reemplazo de Texto**: El script lee dos archivos de Excel que contienen listas de nombres y encuentra los nombres comunes entre ellos. Luego actualiza una capa de texto en un documento de Photoshop con estos nombres.

2. **Exportación de Imágenes por Lote**: Después de actualizar las capas de texto, el script exporta imágenes individualizadas para cada nombre en una carpeta especificada.

## Requisitos
- Python 3.x
- Biblioteca `win32com.client` para automatización en Windows.
- Biblioteca `pandas` para trabajar con archivos de Excel.
- Adobe Photoshop instalado en tu sistema.

## Instalación
1. Instala Python: [https://www.python.org/downloads/](https://www.python.org/downloads/)
2. Instala las bibliotecas de Python necesarias:
   ```bash
   pip install pandas
   pip install pywin32
   ```

## Uso
1. Modifica el script para especificar las rutas de tus archivos de Excel y el documento de Photoshop.
2. Ejecuta el script con el siguiente comando:
   ```bash
   python photoshop_driver.py
   ```
3. Verifica la carpeta de exportación especificada para las imágenes generadas.

## Configuración
- **Archivos de Excel**: Establece las rutas de tus archivos de Excel en las variables `asistencia1` y `asistencia2`.
- **Documento de Photoshop**: Especifica la ruta de tu documento de Photoshop en el método `psApp.Open()`.
- **Carpeta de Exportación**: Establece la carpeta de exportación deseada en la variable `pngfile`.

## Nota Importante
Asegúrate de que Photoshop esté instalado en tu máquina y que hayas guardado tu documento de Photoshop con una capa de texto llamada "TextoEditable".


## Autor
Diego Armando Moreno Peña

Siéntete libre de personalizar este código según tus necesidades específicas.
