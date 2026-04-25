# Excel-To-Csv
- [Description](#description)
- [API Overview](#api-overview)
- [Quick Start](#quick-start)
- [Other Info](#other-info)
- [Known Cell Data Issues](#known-cell-data-issues)
- [Alternate Implementations](#alternate-implementations)
- [Future Work and Non-Goals](#future-work-and-non-goals)
- [Final Thoughts](#final-thoughts)

## Description
Simple tool to convert an Excel worksheet into CSV format.

Implemented using the [Apache POI](https://poi.apache.org/) libraries

> ⚠️ **Warning**
> > Curreently giving the code an 'overhaul'.

## API Overview

The API follows a simple pipeline:
1. Create an `ExcelSheetReadRequest` (defines what to read)
2. Pass it to an `ExcelProcessor` (reads the Excel data)
3. Receive a `SheetContent` result (holds the sheet data)
4. Output it using `CsvWriter` (writes CSV files)

## Quick Start
```java
ExcelSheetReadRequest request = ExcelSheetReadRequest
    .from(Paths.get("input.xlsx"))
    .byName("Sheet1")
    .build();

ExcelProcessor processor = ExcelProcessor.builder()
    .skipBlankRows(true)
    .build();

SheetContent content = processor.readSheets(request).get(0);
CsvWriter.writeToFile(Paths.get("output.csv"), content, QuoteMode.MINIMAL);
```
<details>
  <summary><strong>ExcelSheetReadRequest</strong></summary>

Builds the request and selects which sheet(s) to read.

- `from(Path)`, `from(File)`, `from(URL)`
- `byIndex(...)`, `byIndexes(...)`
- `byName(...)`, `byNames(...)`
- `allSheets()`
- `sheetSelector(...)`
- `password(...)`

</details>

<details>
  <summary><strong>ExcelProcessor</strong></summary>

Reads sheet data and returns `SheetContent`.

- `builder()`
- `readSheets(ExcelSheetReadRequest)`
- `useAdvancedReader(boolean)`

Common options:

- `quoteMode`
- `trimStringValues`
- `skipBlankRows`
- `skipBlankColumns`
- `skipInvisibleCells`
- `saveUnicodeFileWithBom`
- `sanitizeSpaces`
- `sanitizeQuotes`
- `sanitizeDashes`
- `sanitizeDiacritics`

</details>

<details>
  <summary><strong>SheetContent</strong></summary>

Represents a single sheet result.

- `getSheetName()`
- `getMatrix()`
- `getRows()`

</details>

<details>
  <summary><strong>CsvWriter</strong></summary>

Writes `SheetContent` to CSV.

- `writeToFile(Path, SheetContent)`
- `writeToFile(Path, SheetContent, QuoteMode)`
- `writeToDirectory(Path, List<SheetContent>)`
- `writeToDirectory(Path, List<SheetContent>, QuoteMode)`
- `toCsv(SheetContent)`

</details>

<details>
  <summary><strong>QuoteMode</strong></summary>

- `ALWAYS`
- `NORMAL`
- `MINIMAL`
- `NEVER`
</details>

## Other Info
* All rows in the output CSV will have the exact same number of columns. (which will be the maximum non-blank column detected)
* The returned CSV data values are WYSIWYG and should retain the same 'formatting' as the original Excel file. (i.e., Dates and Numeric values)
    * No _formulas_ are copied.  Only the value as it 'physically appears' in a given cell
    * _**(see 'Known Cell Data Issues' for exceptions)_
* Currently, no quotes will be added around 'blank' values
* Empty cells will be converted to empty string (not 'null')

## Known Cell Data Issues
Known Cell Data Formatting Issues include (but are not limited to) the following:
<details>
  <summary>(Click To Expand...)</summary>

Note that some issues below seem to be related to how the Excel file was saved.
Opening the Excel file and then 'resaving' can sometimes resolve issues.

Most cases below appear to be pretty rare (subjectively)

* Sometimes simple numeric values will be returned in a decimal format (and vice versa).
    * i.e. expected "7" but got "7.0" or expected "7.00" but got "7"
* Sometimes 'zero' and 'blank' can get mixed up.
    * i.e. expected "" but got "0.0"
* Advanced parser can throw an exception if encounters a cell without a CellReference
    * a poi-examples class [XLSX2CSV.java](https://github.com/apache/poi/blob/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java) shows a proposed solution, but it only works in a handful of cases.
* Certain Linked or Embedded Objects typically render as "#VALUE!" (Pictures, Stock, Geography, etc.)
* Special or Custom Formats may sometimes render incorrectly.
    * most often noticed with date, time, or numeric values.
    * the ;;; format currently will _NOT_ produce a blank value
* Number precision can sometimes be off (appears to be rare).
    * Examples:
        * "33.8192973" vs "33.8192974"
        * "0.1245" vs "0.124500000"
        * "0.29999" vs "0.2999900001"
* Cells of type DataBar or IconSet will show a value, even if marked as "icon only"

Also note the following _IS_ 'Expected Behavior'
* Cells with 'error values' (#NAME?, #VALUE!, etc.) will appear as such in the output CSV file
</details>

## Alternate Implementations
Searching on the web can yield alternate solutions that require less code.  However, they seem to usually not handle "large" Excel files or don't always handle Blank rows and columns very well

<details>
  <summary>Example Alternate Implementation 1... (Click To Expand)</summary>

An example of a simpler way to read an Excel file without the extra code in this project is below:<br><br>
Additional explanations about the code can be found in [SimplePoiExampleExcelReader.java](src/test/java/com/github/bradjacobs/excel/demo/SimplePoiExampleExcelReader.java)
```java
public List<List<String>> readBasicSheet(Path excelFile) throws IOException {
  DataFormatter formatter = new DataFormatter(true);
  formatter.setUseCachedValuesForFormulaCells(true);

  try (Workbook workbook = WorkbookFactory.create(excelFile.toFile())) {
    Sheet sheet = workbook.getSheetAt(0); // read first sheet
    return IntStream.rangeClosed(0, sheet.getLastRowNum())
            .mapToObj(rowIndex -> {
              Row row = sheet.getRow(rowIndex);
              if (row == null) {
                return List.<String>of();
              }
              return IntStream.range(0, row.getLastCellNum())
                      .mapToObj(colIndex ->
                              formatter.formatCellValue(row.getCell(colIndex)).trim())
                      .collect(Collectors.toList());
            })
            .collect(Collectors.toList());
  }
}
```
</details>
<details>
  <summary>Example Alternate Implementation 2... (Click To Expand)</summary>

From a [StackOverflow Post](https://stackoverflow.com/questions/40283179/how-to-convert-xlsx-file-to-csv), [OrangeDog](https://stackoverflow.com/users/476716/orangedog) points out there is an easier way to get CSV text, which would look something like this:
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
This is a nice solution with _A LOT_ less code.  _**BUT**_... it seems to expose a few limitations with the POI functionality.

Namely:
* empty cells could cause data to seemingly 'shift' to a different column
    * i.e. if no value in Column A, but there is a value in Column B, then the Column B value will show up as the first value in the row.
* Bigger Excel files (> 1 MB ?) will throw an exception with the message: _"The text would exceed the max allowed overall size of extracted text"_
* It will give data from all sheets (even if you only want one)
* The output csv text might not have the cells quoted the way you want (subjective)
</details>

## Future Work and Non-Goals
Possible work items that I _MIGHT_ get around to "eventually" (perhaps)

<details>
  <summary>Todo Item List... (Click To Expand)</summary>

Features:
* Add in addtional 'row and column filtering' (low priority)
  * this would expand on skipping blanks rows/columns.  (i.e. select only certain columns want returned)
  * for column, would select either column index or name

Housekeeping:
* Update/Fix this README with some API documentation
  * CSVWriter usage and parameters
  * Advanced vs Standard mode
* Miscellaneous cleanup and refactoring (ongoing)
* Integrate a real logger into the code
* General Unittest cleanup and add more tests (ongoing)
* Check and fix any circular package dependencies
* Redo the examples
* More Javadocs
* Further updates for API Documentation and README updates.
* Address any of the "Known Cell Data Issues" (above) if possible
* Reorganize Excel Test data for Junit tests.

Other Project Stuff:
* The pom.xml could use some cleanup and organization.
* Put a more legitimate project version in the pom.xml
* Consider making a 'release version' or something that can be referenced via maven dependency
    * need to update groupId and package names from 'com.github...' to 'io.github...'
</details>

Certain items that I _WILL NOT_ get around to.
<details>
  <summary>Won't Fix... (Click To Expand)</summary>

* XLSB support (Excel 2016+)
  * looks like a pain to implement, especially for a format i've rarely come across.
* Date/Time formatting 
  * too many considerations of format, timezone shifting, etc
* Numeric formatting 
  * Do not want to deal with things like: 
    * 2 vs 2.0
    * if 98% should be 0.98 or 98
    * international formats
    * super large, super small numbers
    * etc
  * Harder to deal with in the advaenced event implementation.
* Multiple sheets to SINGLE CSV file.
  * too many concerns if sheets have a different structure.
  * Easy for someone to do programmatically (just merge the SheetContent List<List<String>> values).
* Read all Excel files in a directory.
  * Not hard to write for those who would want that functionality.
* Formula handling Option.
* Reading Cell Comment values.
* Continue-on-error option.
* Allowing for jagged row output.
* CLI support
  * there are probably existing Python versions that can do this better.
</details>

## Final Thoughts
I don't actively work on this project much and only make occasional tweaks just for fun.

This project was originally created in a day, so I'm sure there are specific cases I've missed.  :-)
