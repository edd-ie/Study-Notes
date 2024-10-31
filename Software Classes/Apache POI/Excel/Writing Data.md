#java #excel #microsoft #apachePOI #maven 

# Flow of Operations
1. Create a `workbook`
2. Create a `sheet` in the workbook
3. Create `rows` in the sheet
4. Create `cells` in the rows
5. Input `data` in the cells

Classes used to work with excel files

| **Interface** | **xls class** | **xlsx class** |
| ------------- | ------------- | -------------- |
| Workbook      | HSSFWorkbook  | XSSFWorkbook   |
| Sheet         | HSSFSheet     | XSSFSheet      |
| Row           | HSSFRow       | XSSFRow        |
| Cell          | HSSFCell      | XSSFCell       |
# Creation of workbook
Depending on the version of excel file you want create use the appropriate file
- *Note - xlsx is the modern format for excel files* 

```java
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

XSSFWorkbook workbook = new XSSFWorkbook();
```

# Creation of a sheet
Insert a string as sheet name:
```java
import org.apache.poi.xssf.usermodel.XSSFSheet;

String sheetName = "Sheet1"
XSSFSheet sheet = workbook.createSheet(sheetName);
```

# Prepare data
Store data to be inserted in a 2D data structure like `Object data[][]` or `ArrayList<ArrayList<Object>> data` using `Object` as the data type this will allow as save multiple data types at once due to *polymorphism* .
- Remember to use *wrapper classes* for primitive data types

```java
Object data[][] = {{"Item", "Quantity", "Price"},
				   {"Saws", 10, 9.03},
				   {"Hammer", 3, 12.45}
				   };
```

# Inserting data
Get number of rows and cell using the data structure properties.
Insert the data to the sheet using:
- A for loop if all rows have same number of columns
- An enhanced for loop if some rows have more data than others.
Check the type of value that is to be inserted and cast it to the appropriate data type

**For Loop** 
```java
import org.apache.poi.xssf.usermodel.XSSFRow;

int rows = data.length;
int colls = data[0].length

for (int i = 0; i < rows; i++) {  
    XSSFRow row = sheet.createRow(i);  
    for (int j = 0; j < cols; j++) {  
        Object val = data.get(i).get(j);  
        switch (val.getClass().getSimpleName()) {  
            case "String":  
                row.createCell(j).setCellValue((String) val);  
                break;  
            case "Integer":  
                row.createCell(j).setCellValue((Integer) val);  
                break;  
            case "Double":  
                row.createCell(j).setCellValue((Double) val);  
                break;  
            case "Boolean":  
                row.createCell(j).setCellValue((Boolean) val);  
                break;  
        }  
    }  
}
```

**Enhance For loop ** 
```java
import org.apache.poi.xssf.usermodel.XSSFRow;

int rows = 0;  
int cols;  
  
for(ArrayList<Object> row : data){  
    XSSFRow xssfRow = sheet.createRow(rows++);  
    cols = 0;  
    for(Object cell : row){  
        switch (cell.getClass().getSimpleName()) {  
            case "String":  
                xssfRow.createCell(cols++).setCellValue((String) cell);  
                break;  
            case "Integer":  
                xssfRow.createCell(cols++).setCellValue((Integer) cell);  
                break;  
            case "Double":  
                xssfRow.createCell(cols++).setCellValue((Double) cell);  
                break;  
            case "Boolean":  
                xssfRow.createCell(cols++).setCellValue((Boolean) cell);  
                break;  
        }  
    }  
}
```

# Saving the workbook
To save the workbook:
- Create a new file in write mode using `FileOutputStream`
- Ensure to place it in a *try-catch()* as it throw and `IOException`
- Write the workbook to the file.
- Ensure to use robust error handling.

```java
import java.util.logging.Level;  
import java.util.logging.FileHandler;  
import java.util.logging.Logger;
import java.io.FileOutputStream;  
import java.io.IOException;

// Implementation to save the workbook to the specified file path  
FileOutputStream fileOut = null;  
String Filepath = "/Folde/Name.xlsx"
Logger logger = Logger.getLogger(FileHandler.class.getName());
  
try {  
    fileOut = new FileOutputStream(filePath);  
    workbook.write(fileOut);  
    logger.log(Level.INFO, "File saved successfully." + filePath);  
} catch (IOException e) {  
    logger.log(Level.SEVERE, "File Could not be saved." + filePath, e);  
} finally {  
    try {  
        if (fileOut != null) {  
            fileOut.close();  
            workbook.close();  
            logger.log(Level.INFO, "File closed successfully." + filePath);  
        }  
        workbook.close();  
    } catch (Exception e) {  
        logger.log(Level.SEVERE, "Error closing the file" + filePath, e);  
    }  
}
```