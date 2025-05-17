# Excel-To-Csv
- [Description](#Description)
- [Examples](#Examples)
  * [Basic](#Basic)
  * [Advanced](#Advanced)
- [Usage](#Usage)
  * [Overview](#Overview)
  * [ExcelReaderDetails](#ExcelReaderDetails)
  + [BuilderDetails](#BuilderDetails)
- [OtherInfo](#OtherInfo)
- [TODOs](#TODOs)
- [AlternateImplementations](#AlternateImplementations)
- [FinalThoughts](#FinalThoughts)


## Description
Simple tool to convert an Excel worksheet into CSV format.

Implemented using the [Apache POI](https://poi.apache.org/) libraries

## Examples

### Basic
```java
// read excel worksheet and write output to a file
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(Paths.get("input.xlsx"), Paths.get("output.csv"));
// or
excelReader.convertToCsvFile(new File("input.xlsx"), new File("output.csv"));
```

```java
// get a single string representing the entire worksheet in CSV format
ExcelReader excelReader = ExcelReader.builder().build();
String csvText = excelReader.convertToCsvText(Paths.get("input.xlsx"));
// or
String csvText = excelReader.convertToCsvText(new File("input.xlsx"));
```

```java
// get 2-D string array representing the entire worksheet (each value represents a 'cell')
ExcelReader excelReader = ExcelReader.builder().build();
String[][] csvData = excelReader.convertToDataMatrix(Paths.get("input.xlsx"));
// or
String[][] csvData = excelReader.convertToDataMatrix(new File("input.xlsx"));
```

### Advanced
```java
// write csv file w/ specific settings
ExcelReader excelReader = ExcelReader.builder()
        .quoteMode(QuoteMode.LENIENT) // only quote values if necessary
        .sheetIndex(1) // grab the 2nd worksheet
        .skipEmptyRows(true) // ignore any empty rows from the Excel worksheet
        .build();
excelReader.convertToCsvFile(new File("input.xlsx"), new File("output.csv"));
```
```java
// fetch Excel file from external URL location and save as a local csv file.
ExcelReader excelReader = ExcelReader.builder().build();
excelReader.convertToCsvFile(new URL("https://some.domain.com/input.xlsx"), new File("output.csv"));
```

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
| FIELD                  | REQUIRED | DEFAULT         | DETAILS                                                                                                                                                                                                                                                                        |
|------------------------|----------|-----------------|--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------|
| quoteMode              | NO       | NORMAL          | how aggressive to wrap quotes around values<br><br>*ALWAYS*: always put quotes around values<br>*NORMAL*: put quotes around most values that are non-alphanumeric<br>*LENIENT*: only add quotes around values that are needed to be CSV compliant<br>*NEVER*: never add quotes |
| sheetIndex             | NO       | 0               | 0-based index of which worksheet to convert to CSV                                                                                                                                                                                                                             |
| sheetName              | NO       | (blank)         | Name of the worksheet tab to be converted to CSV<br> (if set then 'sheetIndex' is ignored)                                                                                                                                                                                     |
| skipEmptyRows          | NO       | false           | filter out all 'blank' rows from the Excel worksheet                                                                                                                                                                                                                           |
| saveUnicodeFileWithBom | NO       | true            | prepend 'BOM' to output CSV file if unicode characters were detected.                                                                                                                                                                                                          |
| sanitizeWhitespace     | NO       | true            | replace any unicode or abnormal space character (i.e. nbsp) with a normal space                                                                                                                                                                                                |
| sanitizeQuotes         | NO       | true            | replace any special single/double quotes (i.e. smart quotes) with normal quotes                                                                                                                                                                                                |
| sanitizeDashes         | NO       | false           | replace any specia dash/hyphen character (i.e. em dash) with normal dash                                                                                                                                                                                                       |
| sanitizeDiacritics     | NO       | false           | replace diacritic characters with its basic counterpart (i.e. 'é' -> 'e', 'Ç' -> 'C')                                                                                                                                                                                          |

## OtherInfo
* All rows in the output CSV will have the exact same number of columns. (which will be maximum non-blank column detected)
* The CSV data values should retain same 'formatting' as the original Excel file. (i.e. Dates and Numeric values)
* No _formulas_ are copied.  Only the value as it 'physically appears' in a given cell
  * (_Exception_ - cells that have a conditional formatting rule, such as a "data bar", may result in incorrect or invalid cell value)
* Currently no quotes will be added around 'blank' values 
* Empty cells will be converted to empty string (not 'null')
* All cell values are "trimmed" (assuming one usually does NOT want leading/trailing whitespace)

## TODOs
A work item list that I might get around to "eventually" (perhaps)
* Put a more legitimate project version in the pom.xml
* Consider making a 'release version' or something that can be referenced via maven dependency
* Integrate a real logger into the code.
* Research better handling for cells that show a "data bar or other custom formatting". (see comment in previous section)
* Add more JavaDocs
* More Unittest cleanup


## AlternateImplementations
Searching on the web can yield alternate solutions that require less code.  However, they seem to usually not handle "large" Excel files or doesn't always handle Blank rows and columns very well

<details>
  <summary>Example Alternate Implementation... (Click Me)</summary>

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
This appears to work with a lot less code.  _**BUT**_... it seems to expose a few limitations with the POI functionality.

Namely:
* empty cells could cause data to seemingly 'shift' to a different column
  * i.e. if no value in Column A, but is a value in Column B, then the Column B value will show up as the first value in the row.
* Bigger Excel files (> 1MB ?) will throw an exception with message: _"The text would exceed the max allowed overall size of extracted text"_
* It will give data from all sheets (even if you only want one)
* The output csv text might not have the cells quoted the way you want (subjective)
</details>

## FinalThoughts
I don't actively work on this project much and only make occasional tweaks just for fun.

This project was originally created in a day, so I'm sure there are specific cases I've missed.  :-) 
