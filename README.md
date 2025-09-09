# matching_tables_and_PB
A script to match a inventory of assets with results PB

# Explicación del Script 📄

## En Español 🇪🇸

Este script en Python sirve para procesar datos de inventario y resultados dentro de archivos Excel, y aplicar formatos condicionales para facilitar la identificación de ciertos valores. 

### Funcionalidades principales:

- Carga dos archivos Excel (inventario y resultados) junto con sus hojas respectivas.  
- Limpia y normaliza valores en columnas específicas (ejemplo: columnas D y E).  
- Detecta valores duplicados en la columna D del inventario.  
- Compara valores del inventario con la lista de resultados para identificar coincidencias.  
- Abre el archivo de inventario para modificarlo visualmente:  
  - Aplica color azul y fuente Calibri a las celdas que coinciden con resultados o que están duplicadas (columna D).  
- Busca dinámicamente la columna con encabezado `"Executed"` y, para las filas que contienen celdas marcadas en D o E, pone el valor `"Ejecutado"` en esa columna, aplicando el mismo formato.  
- Guarda los cambios realizados en el archivo de inventario.  
- Controla errores e informa si algo falla durante la ejecución.  

Este proceso ayuda a visualizar rápidamente qué elementos del inventario fueron encontrados en la lista de resultados y cuáles están duplicados, facilitando auditorías y seguimientos. ✅

---

## In English 🇬🇧

This Python script processes inventory and results data from Excel files and applies conditional formatting to facilitate data analysis and visualization.

### Main functionalities:

- Loads two Excel files (inventory and results) with their respective sheets.  
- Cleans and normalizes values in specific columns (example: columns D and E).  
- Detects duplicated values in column D of the inventory.  
- Compares inventory values against the results list to find matches.  
- Opens the inventory file to visually modify it:  
  - Applies blue fill color and Calibri font to cells matching results or duplicated (column D).  
- Dynamically finds the column with header `"Executed"` and for rows with marked cells in columns D or E, sets the value `"Executed"` in that column, applying the same format.  
- Saves the changes back into the inventory file.  
- Handles errors and informs if something goes wrong during execution.  

This script helps quickly visualize which inventory items were found in the results list and which are duplicated, assisting audits and tracking. ✅
