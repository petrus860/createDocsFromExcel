# createDocsFromExcel
this script read a excel sheet and relate the columns and generate docx files with them


Requisitos para que funcione este script
instalar Python 3.9.5
	instalar librerias
		pip install openpyxl
		pip install python-docx
		
		
 solucion, error --> raise ValueError('Max value is {0}'.format(self.max)) ValueError: Max value is 52
he buscado en los xml el valor 66 y le he puesto 50
https://stackoverflow.com/questions/50236928/openpyxl-valueerror-max-value-is-14-when-using-load-workbook

Ahora mismo estoy usando el fichero "Sample_Model0001.xlsx" que funciona correctamente.

El script funciona de la siguiente manera:
        Leer las columnas ABCDFGH y mete los datos en una listas
        Recorre las listas de cada columna y prepara los datos para crear un archivo doc por cada caso de prueba con sus correspondientes pasos de prueba.
        Con los datos crear un archivo en formato .doc con un formato especifico 

Requirements
	openpyxl~=3.0.8
	docx~=0.2.4
