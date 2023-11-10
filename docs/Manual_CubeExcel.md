



# Cube Excel
  
Module to work with Excel OLAP/Cube Pivot Tables  

*Read this in other languages: [English](Manual_CubeExcel.md), [Português](Manual_CubeExcel.pr.md), [Español](Manual_CubeExcel.es.md)*
  
![banner](imgs/Banner_CubeExcel.jpg)
## How to install this module
  
To install the module in Rocketbot Studio, it can be done in two ways:
1. Manual: __Download__ the .zip file and unzip it in the modules folder. The folder name must be the same as the module and inside it must have the following files and folders: \__init__.py, package.json, docs, example and libs. If you have the application open, refresh your browser to be able to use the new module.
2. Automatic: When entering Rocketbot Studio on the right margin you will find the **Addons** section, select **Install Mods**, search for the desired module and press install.  


## Description of the commands

### Refresh Pivot table
  
Refresh a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivot table name|Name: |

### Add Field - Cube
  
Add field to a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivot table name|Name: |
|Select option|Select what you want to add||
|Field to add |Name of the field to add|[Field].[Subfield]: |

### Remove Field - Cube
  
Remove field from a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivot table name|Name: |
|Field to add |Name of the field to add|[Field].[Subfield]: |

### Add Filter - Cube
  
Filter a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Filter |Filter to apply|[Field].[SubField]: |
|Opened Filter |Opened Filter|[filter].[All filter].[data].[more data] |

### Apply Filter - Cube
  
Filter a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Filter |Filter to apply|[Field].[SubField]: |
|Opened Filter |Opened Filter|Filter to apply |
|Number as text||False|

### Clear Filter - Cube
  
Clear a filter from a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivote table name|Name: |
|Filter |Filter to apply|[Field].[SubField]: |

### Clear All Filters - Cube
  
Clear all filters from a pivot table
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Name of the sheet where the pivot table is located|Sheet1|
|Pivote table name |Pivote table name|Name: |

### List Visible Items - Cube
  
List visible items of a field from a Pivot Table - Cube
|Parameters|Description|example|
| --- | --- | --- |
|Sheet |Sheet name where the field is located|Sheet1|
|Pivote table name |Pivote table name where the field is located|Name: |
|Field |Field name|[Field]: |
|Assign result to variable |Variable name where the result will be stored|Variable|

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
