# Cube Excel
  
Module with options to work with pivot tables in excel cube files  

*Read this in other languages: [English](Manual_CubeExcel.md), [Español](Manual_CubeExcel.es.md).*
  
![banner](imgs/Banner_CubeExcel.png)
## How to install this module
  
__Download__ and __install__ the content in 'modules' folder in Rocketbot path  



## Description of the commands

### Refresh Pivot table
  
Refresh a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivot table name|Name: |

### Add field - Cube
  
Add field to a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivot table name|Name: |
|Select option|Select what you want to add||
|Field to add |Name of the field to add|[Field].[Subfield]: |

### Add Filter - Cube
  
Filter a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Filter |Filter to apply|[Field].[SubField]: |
|Opened Filter |Opened Filter|[filter].[All filter].[data].[more data] |

### Open Field
  
Open a field. Similar to click on plus button
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Field |Field name|[Field].[type]: |
|Cell field |Cell field name|[Field].[All Field].[cell name].[other cell name]: " |

### List Fields
  
List all available table fields
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Assign result to variable |Variable name where the result will be stored|Variable|

### List SubField
  
List Pivot Table from a Cube Field
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the field is located|Sheet1|
|Pivote table name |Pivote table name where the field is located|Name: |
|Field |Field name|[Field]: |
|Assign result to variable |Variable name where the result will be stored|Variable|

