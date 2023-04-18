
# Cargar las librerías necesarias
library(officer)
library(readxl)

# Ruta del archivo de Word con el formato previamente diseñado
ruta_documento <- "CAMBIO DE NOTA.docx"

# Ruta del archivo de Excel con la lista de nombres y códigos
ruta_excel <- "LISTA.xlsx"

# Leer el archivo de Excel
datos_excel <- read_excel(ruta_excel)

# Obtener la lista de nombres y códigos del archivo de Excel
nombres <- datos_excel$Nombre
codigos <- as.character(datos_excel$codigo)

# Loop a través de los nombres y códigos para buscar y reemplazar el contenido en el documento de Word
for (i in 1:length(nombres)) {
  # Cargar el documento de Word
   doc <- read_docx(ruta_documento)
   # Crear una copia del documento para cada nombre y código
  doc_copia <- doc
  
  # Buscar y reemplazar el nombre en el documento de Word
  doc_copia <- body_replace_all_text(doc_copia, old_value = "NOMBRE_ESTUDIANTE", new_value = nombres[i])
  
  # Buscar y reemplazar el código en el documento de Word
  doc_copia <- body_replace_all_text(doc_copia, old_value = "CODIGO_ESTUDIANTE", new_value = codigos[i])
  
  # Crear un nuevo documento de Word con el contenido modificado
  nombre_documento <- paste("Documento_", nombres[i], ".docx", sep = "")
  print(nombre_documento)
  print(paste("NOMBRE:", nombres[i], "CODIGO:", codigos[i]))
  print("------------")
  print(doc, target = nombre_documento)
  #print(doc_copia, target = nombre_documento)
}
