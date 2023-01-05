# Cube Excel
  
Módulo con opciones para trabajar tablas dinámicas en archivos cube de excel  

*Read this in other languages: [English](Manual_CubeExcel.md), [Español](Manual_CubeExcel.es.md).*
  
![banner](imgs/Banner_CubeExcel.png)
## Como instalar este módulo
  
__Descarga__ e __instala__ el contenido en la carpeta 'modules' en la ruta de Rocketbot.  



## Descripción de los comandos

### Actualizar tabla dinámica
  
Actualiza una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |

### Agregar campo - Cube
  
Agrega un campo a una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Selecciona una opción|Selecciona lo que deseas agregar||
|Campo a agregar|Nombre del campo a agregar|[Field].[Subfield]: |

### Agregar filtro - Cube
  
Filtra una tabla dinámica
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra la tabla dinámica|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Filter |Filtro a aplicar|[Field].[SubField]: |
|Filtro abierto |Filtro abierto|[filter].[All filter].[data].[more data]: |

### Abrir Campo
  
Abre un campo. Similar a hacer click en botón más
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Campo |Nombre del campo|[Field].[type]: |
|Campo celda |Nombre del campo de celda|[Field].[All Field].[cell name].[other cell name]: |

### Listar Campos
  
Lista todos los campos disponibles
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica|Name: |
|Asignar resultado a variable|Nombre de la variable donde se almacenará el resultado|Variable|

### Listar SubCampo
  
Lista todos los campos dinámicos de un campo Cube
|Parámetros|Descripción|ejemplo|
| --- | --- | --- |
|Hoja |Nombre de la hoja donde se encuentra el campo|Hoja 1|
|Nombre de la tabla dinámica |Nombre de la tabla dinámica donde se encuentra el campo|Name: |
|Campo |Nombre del campo|[Field]: |
|Asignar resultado a variable|Nombre de la variable donde se almacenará el resultado|Variable|
