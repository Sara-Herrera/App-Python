# App-Python

Esta aplicación permite generar informes sobre información de variantes descargadas de ClinVar, clasificándolas según su clasificación (ejemplo: https://www.ncbi.nlm.nih.gov/clinvar/?term=SMN1%5Bgene%5D&redir=gene). Las categorías incluidas son: *"Conflicting classifications", "Benign", "Likely benign", "Uncertain significance", "Likely pathogenic"* y *"Pathogenic"*.

### Funcionalidades
1.	Generación de Informes: Clasificación de variantes según su clasificación clínica.
2.	Añadir Información de Estudio: Permite añadir información adicional de estudios. Es posible crear un archivo que contenga datos de varios estudios, los cuales se pueden exportar y cargar al momento de generar el informe.
   
#### Contenido
- Proyecto de Python: Contiene el código fuente y archivos necesarios:
  - `code.py`: Archivo principal del código fuente.
  - `config.json`: Archivo de configuración.
  - `labels.py`: Archivo de etiquetas.
  - `images_folder`: Carpeta de imágenes.

- Inputs:
  - Ejemplo de archivo de análisis.
  - Ejemplo de archivo de datos de estudio.

- Output:
  - Ejemplo de informe generado.

