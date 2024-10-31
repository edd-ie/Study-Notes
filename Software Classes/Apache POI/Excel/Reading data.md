#java #excel #microsoft #apachePOI #maven
# Introduction
Apache POI has 4 interface to work with 2 kinds of excel file `xls` and `xlsx` 

| **Interface** | **xls class** | **xlsx class** |
| ------------- | ------------- | -------------- |
| Workbook      | HSSFWorkbook  | XSSFWorkbook   |
| Sheet         | HSSFSheet     | XSSFSheet      |
| Row           | HSSFRow       | XSSFRow        |
| Cell          | HSSFCell      | XSSFCell       |

## [Download & Setup]()

*[Video link](https://www.youtube.com/watch?v=sg8_jUcqQaA)*
To set up your own library: [Download File](https://poi.apache.org/download.html#POI-5.3.0)

To setup using maven get the [Apache POI Common](https://mvnrepository.com/artifact/org.apache.poi/poi) and  [Apache POI API Based On OPC and OOXML Schemas](https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml) the [maven repository]([Maven Repository: apache poi](https://mvnrepository.com/search?q=apache+poi))

Or paste this in `POM.xml` for version `5.3.0`  in the `<dependecies>` :
```xml
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi -->  
<dependency>  
    <groupId>org.apache.poi</groupId>  
    <artifactId>poi</artifactId>  
    <version>5.3.0</version>  
</dependency>  
  
<!-- https://mvnrepository.com/artifact/org.apache.poi/poi-ooxml -->  
<dependency>  
    <groupId>org.apache.poi</groupId>  
    <artifactId>poi-ooxml</artifactId>  
    <version>5.3.0</version>  
</dependency>
```

# Read files
1. Get file path.
2. Open the file in read mode using `FileInputStream` class
```java
import java.io.FileInputStream;

String filePath = "folde/file.xlsx";
FileInputStream file = null;

try {  
    file = new FileInputStream(filePath);  
} catch (Exception e) {  
    e.printStackTrace();  
}

/** For more robust logging of the error:**/
import java.io.FileInputStream;  
import java.io.FileNotFoundException;  
import java.util.logging.FileHandler;  
import java.util.logging.Level;  
import java.util.logging.Logger;

private static final Logger logger = Logger.getLogger(FileHandler.class.getName());

private FileInputStream file = null;
String filePath = "folde/file.xlsx";

try {  
    file = new FileInputStream(filePath);  
} catch (FileNotFoundException e) {  
    logger.log(Level.SEVERE, "File not found: " + filePath, e);  
}

```

3. Get the workbook from the file.
```java 
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public XSSFWorkbook readWorkbook(FileInputStream file) throws IOException {  
    // Implementation to read Excel file  
    return new XSSFWorkbook(file);  
}
```

4. Access a sheet:
```java
import org.apache.poi.xssf.usermodel.XSSFSheet; 

int sheetNumber = 1;
XSSFSheet sheet = workbook.getSheetAt(sheetNumber);  
```

5. To get the number of `rows` in a sheet:
```java
int row =  sheet.getLastRowNum();
```

6. To get the a `row` in a sheet:
```java
import org.apache.poi.xssf.usermodel.XSSFRow;

int indexOfRow = 1;
XSSFRow row =  sheet.getRow(indexOfRow);
```

7. To get the number of columns / `cells`  in a row:
```java
int indexOfRow = 1;
int cols =  sheet.getRow(indexOfRow).getLastCellNum();
```

8. To get the a `value` in a cell:
```java
import org.apache.poi.xssf.usermodel.XSSFCell;

XSSFCell value =  sheet.getRow(1).getCell(0);
```

9. Reading `entire sheet` :
```java
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFCell;

int rows =  sheet.getLastRowNum();
int cols =  sheet.getRow(1).getLastCellNum();

XSSFRow line;
XSSFCell cell;

for(int x = 0; x < rows; x++){
	line = sheet.getRow(x);

	for(int y = 0; y < cols; y++){
		cell = line.getCell(y);
	}
}
```

10. To get the `data type` of data save in cell:
```java
cell.getCellType();
```

11. Getting the data by data type:
```java
XSSFCell cell = row.getCell(2);

switch(cell.getCellType()){
	case STRING: cell.getStringCellValue(); break;
	case NUMERIC: cell.getNumericCellValue(); break;
	case BOOLEAN: cell.getBooleanCellValue(); break;
}
```

12. Reading data using an iterator:
```java
Iterator moveDown = sheet.iterator();
Iterator moveLeft;

XSSFRow row;
XSSFCell cell;

while(move.hasNext()){
	row = (XSSFRow)moveDown.next();
	moveLeft = row.cellIterator();

	while(moveLeft.hasNext()){
		cell = (XSSFCell)moveLeft.next();

		switch(cell.getCellType()){
			case STRING: cell.getStringCellValue(); break;
			case NUMERIC: cell.getNumericCellValue(); break;
			case BOOLEAN: cell.getBooleanCellValue(); break;
		}
	}
}
```