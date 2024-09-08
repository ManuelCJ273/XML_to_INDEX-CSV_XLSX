Algoritmo para convertir el XML (De la junta de andalucia de certificados de eficiencia enegetica) en CSV y XLSX (Excel) 

El código en concreto se pasa largo tiempo en función del ordenador que se tenga, revisando los XML, creando una cadena de archivos cada 30.000 registros y un archivo de incidencias con los datos que de alguna manera presentan discrepancias, como por ejemplo códigos postales que no cuadran con el nombre de la provincia, campos vacíos, etc.

Una vez creamos los archivos, se pueden subir los CSV directamente a la base de datos. Yo los he pasado por Excel para comprobar y los he guardado en CSV separados por ";" con codificación UTF-8 para evitar perder registros. Esto no es un manual de SQL, sino una guía para desgranar la montaña de datos de la Junta de Andalucía. Podemos resumir lo siguiente:

Se han procesado: 916.636 certificados

Incidencias: 7.632 (0.83%)

Registros con referencias catastrales de otras comunidades: 34 (0.0037%)

Se ha construido una base de datos en SQL con casi 1.000.000 de registros obtenidos a partir de los XML comprimidos en 7z de la Junta de Andalucía. Las incidencias se refieren al procesar esos datos con campos erróneos, y no se presumen incidencias de cada expediente.

Se ha observado la existencia de registros con referencia catastral fuera de Andalucía, aunque la dirección en la consulta oficial del registro es una dirección de Almería, lo que podría indicar errores en la introducción de la referencia catastral.

Por cierto, estos datos se basan en la actualización del 22 de agosto de 2024, 2:00 (UTC+02:00).

