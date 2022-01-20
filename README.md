# Excel-To-Csv
- [Background](#Background)
- [Usage](#Usage)
  * [Overview](#Overview)
    + [Builder_Details](#Builder_Details)
    + [ExcelReader_Details](#ExcelReader_Details)
- [Examples](#Examples)
  * [Basic](#Basic)
  * [Advanced](#Advanced)
- [Other_Info](#Other_Info)
- [Testing](#Testing)

## Background
Different cases arose where i wanted to parse data out of an Excel file, but I wasn't overly interested in doing special Excel parse logic for each case.  Thus the origin of this
 tool which basically just converts an Excel worksheet into a CSV-format, then can work with CSV formatted data.

## Usage
### Overview
1. Create a new ExcelReader via builder() method.
2. Execute desired methods on ExcelReader

### ExcelReader Details
| METHOD              | INPUTS                       | OUTPUT     | DESCRIPTION                                                                                                                                      |   |
|---------------------|------------------------------|------------|--------------------------------------------------------------------------------------------------------------------------------------------------|---|
| convertToCsvText    | Excel File                   | String     | Given Excel file input return a String representing the Worksheet as CSV                                                                         |   |
| convertToDataMatrix | Excel File                   | String[][] | Given Excel file input return a 2-D String array representing the Worksheet as CSV<br> (each array element represents a cell from the worksheet) |   |
| convertToCsvFile    | Excel File & Output CSV File | (none)     | Given Excel file input write output directly to a specified destination file.  

### Builder Details
| FIELD         | REQUIRED | DEFAULT | DETAILS                                                                                                                                                                                                                                                                                  |   |
|---------------|:--------:|:-------:|------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|---|
| quoteMode     | NO       | NORMAL  | how aggressive / lenient it should wrap quotes around values<br><br>*ALWAYS*: always put quotes around values<br>*NORMAL*: put quotes around most values that are non-alphanumeric<br>*LENIENT*: only add quotes around values that are needed to be CSV compliant<br>*NEVER*: never add quotes |   |
| sheetIndex    | NO       | 0       | 0-based index of which worksheet to convert to CSV                                                                                                                                                                                                                                       |   |
| sheetName     | NO       | (blank) | Name of the worksheet tab to be converted to CSV<br> (if set then 'sheetIndex' is ignored)                                                                                                                                                                                                       |   |
| skipEmptyRows | NO       | false   | if true, then any 'all blank' rows from the Excel worksheet will be ignored.                                                                                                                                                                                                             |   |

## Examples
### Basic
```java
// get a single string representing the entire worksheet in CSV format
File excelFile = new File("/some/path/excelfile.xlsx");
ExcelReader excelReader = ExcelReader.builder().build();
String csvText = excelReader.convertToCsvText(excelFile);
```
```java
// get 2-D string array the entire worksheet in CSV format (each value represents a 'cell')
File excelFile = new File("/some/path/excelfile.xlsx");
ExcelReader excelReader = ExcelReader.builder().build();
String[][] csvData = excelReader.convertToDataMatrix(excelFile);
```
```java
// read excel worksheet and write output to a file
File excelFile = new File("/some/path/excelfile.xlsx");
File outputFile = new File("/different/path/test_data.csv");
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(excelFile, outputFile);
```

### Advanced
```java
// write csv file w/ specific settings
File excelFile = new File("/some/path/excelfile.xlsx");
File outputFile = new File("/different/path/test_data.csv");
ExcelReader excelReader = ExcelReader.builder()
        .setQuoteMode(QuoteMode.LENIENT) // only quote values if necessary
        .setSheetIndex(1) // grab the 2nd worksheet
        .setSkipEmptyRows(true) // ignore any empty rows from the Excel worksheet
        .build();
excelReader.convertToCsvFile(excelFile, outputFile);
```
```java
// fecth Excel file from external URL location and save as a local csv file.
URL excelFileUrl = new URL("https://some.domain.com/download/1/docs/SampleData.xlsx");
File outputFile = new File("/different/path/test_data.csv");
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(excelFileUrl, outputFile);
```

## Other_Info
* All rows in the output CSV will have the exact same number of columns. (which will be max non-blank column detected)
* The CSV data should remain the same as the Excel file 
  * i.e. Dates and Numeric values should retain their existing formatting
* No _formulas_ are copied.  Only the value that appears in a given cell
* Currently no quotes will be added around 'blank' values 
* Empty cells will be converted to empty string (not 'null')

## Testing
The project contains unittest for most of the basic functionality.

However, the following scenarios have **NOT** been tested...
* BIG Excel files
* Files with FTP urls
* Older/newer versions of Excel files.
* Excel files that were originally generated on Windows
* Compatible w/ different JAVA versions.
* Accessing files where don't have the required read and/or write permissions
* Any URLs that require special authentication or headers.
* ".xls" files
* Unicode / extended characters

I originally did this project in a day, so i'm sure there are other items I've probably missed.  ;-) 
