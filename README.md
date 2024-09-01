# Excel-To-Csv
- [Description](#Description)
- [Usage](#Usage)
  * [Overview](#Overview)
    + [BuilderDetails](#BuilderDetails)
    + [ExcelReaderDetails](#ExcelReaderDetails)
- [Examples](#Examples)
  * [Basic](#Basic)
  * [Advanced](#Advanced)
- [OtherInfo](#OtherInfo)
- [Testing](#Testing)
- [TechNotes](#TechNotes)
  * [AlternateImplemenation](#AlternateImplemenation)

## Description
Simple tool to convert an Excel worksheet into CSV format.

Implemented using the [Apache POI](https://poi.apache.org/) libraries


## Usage
### Overview
1. Create a new ExcelReader via builder() method.
2. Execute desired methods on ExcelReader

### ExcelReaderDetails
| METHOD              | INPUTS                       | OUTPUT     | DESCRIPTION                                                                                                                                      |
|---------------------|------------------------------|------------|--------------------------------------------------------------------------------------------------------------------------------------------------|
| convertToCsvText    | Excel File                   | String     | Given Excel file input return a String representing the Worksheet as CSV                                                                         |
| convertToDataMatrix | Excel File                   | String[][] | Given Excel file input return a 2-D String array representing the Worksheet as CSV<br> (each array element represents a cell from the worksheet) |
| convertToCsvFile    | Excel File & Output CSV File | (none)     | Given Excel file input write output directly to a specified destination file.  

### BuilderDetails
| FIELD         | REQUIRED | DEFAULT | DETAILS                                                                                                                                                                                                                                                                                         |
|---------------|:--------:|:-------:|-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| quoteMode     | NO       | NORMAL  | how aggressive / lenient it should wrap quotes around values<br><br>*ALWAYS*: always put quotes around values<br>*NORMAL*: put quotes around most values that are non-alphanumeric<br>*LENIENT*: only add quotes around values that are needed to be CSV compliant<br>*NEVER*: never add quotes |
| sheetIndex    | NO       | 0       | 0-based index of which worksheet to convert to CSV                                                                                                                                                                                                                                              |
| sheetName     | NO       | (blank) | Name of the worksheet tab to be converted to CSV<br> (if set then 'sheetIndex' is ignored)                                                                                                                                                                                                      |
| skipEmptyRows | NO       | true    | if true, then any 'all blank' rows from the Excel worksheet will be ignored.                                                                                                                                                                                                                    |

## Examples
### Basic
```java
// get a single string representing the entire worksheet in CSV format
ExcelReader excelReader = ExcelReader.builder().build();
String csvText = excelReader.convertToCsvText(new File("input.xlsx"));
```
```java
// get 2-D string array the entire worksheet in CSV format (each value represents a 'cell')
ExcelReader excelReader = ExcelReader.builder().build();
String[][] csvData = excelReader.convertToDataMatrix(new File("input.xlsx"));
```
```java
// read excel worksheet and write output to a file
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(new File("input.xlsx"), new File("output.csv"));
```

### Advanced
```java
// write csv file w/ specific settings
ExcelReader excelReader = ExcelReader.builder()
        .setQuoteMode(QuoteMode.LENIENT) // only quote values if necessary
        .setSheetIndex(1) // grab the 2nd worksheet
        .setSkipEmptyRows(true) // ignore any empty rows from the Excel worksheet
        .build();
excelReader.convertToCsvFile(new File("input.xlsx"), new File("output.csv"));
```
```java
// fetch Excel file from external URL location and save as a local csv file.
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(new URL("http://some.domain.com/input.xlsx"), new File("output.csv"));
```

## OtherInfo
* All rows in the output CSV will have the exact same number of columns. (which will be max non-blank column detected)
* The CSV data values should retain same 'formatting' as the original Excel file. (i.e. Dates and Numeric values)
* No _formulas_ are copied.  Only the value as it 'physically appears' in a given cell
* Currently no quotes will be added around 'blank' values 
* Empty cells will be converted to empty string (not 'null')
* All cell values are "trimmed" (assuming one usually does NOT want leading/trailing whitespace)

## Testing
The project contains unittests for most of the basic functionality.

However, the following scenarios have either no testing or very limited testing...
* HUGE Excel files _(limited testing only)_
* Older/Newer versions of Excel files.
* Excel files that were originally generated on Windows
* Compatibility w/ different JAVA versions.
* Accessing files that do not have the necessary read/write permissions
* Unicode / extended characters
* Worksheets containing nested charts.
* Use of the URL input in lieu of File input

## TechNotes
* I don't actively maintain this project, and make occasional tweaks just for fun.
* This project is still compiling with JDK 8.
  * The original thought was in case need to use this code with other libraries using old JDK.  _However_... at this point anything still on JDK 8 seems silly, so i will eventually bump up the JDK version.
* The dependency versions are getting out-of-date.  Will update 'eventually'
* Cannot seem to recall why I used `testng` instead of `junit`.
  * I may switch over the tests to be junit at some point.

### AlternateImplemenation
From a [StackOverflow Post](https://stackoverflow.com/questions/40283179/how-to-convert-xlsx-file-to-csv), [OrangeDog](https://stackoverflow.com/users/476716/orangedog) points out there is an easier way to get CSV text, which would look something ike this:
```java
XSSFWorkbook input = new XSSFWorkbook(new File("input.xlsx"));
try (CSVPrinter output = new CSVPrinter(new FileWriter("output.csv"), CSVFormat.DEFAULT);) {
    String tsv = new XSSFExcelExtractor(input).getText();
    BufferedReader reader = new BufferedReader(new StringReader(tsv));
    reader.lines().map(line -> line.split("\t")).forEach(k -> {
        try { output.printRecord(Arrays.asList(k)); } catch (IOException e) { /* ignore */ }
    });
}
```
This appears to work and requires a lot less code.  _BUT_... it seems to expose a few limitations with the POI functionality.

Namely:
* empty cells could cause data to seemingly 'shift' to a different column
  * i.e. if no value in Column A, but is a value in Column B, then the Colum B value will show up as the first value in the row.
* Bigger Excel files (>1MB ?) will throw an exception with message: _"The text would exceed the max allowed overall size of extracted text"_
* The output csv text might not have the cells quoted the way you want (subjective)


This project was originally created in a day, so i'm sure there are other items I've missed.  ;-) 
