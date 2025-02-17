/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;

class ExcelSheetReader {
    private static final boolean EMULATE_CSV = true;
    private static final DataFormatter EXCEL_DATA_FORMATTER = new DataFormatter(EMULATE_CSV);
    static {
        // set true to get actual visible value in a formula cell, and not the raw formula itself.
        EXCEL_DATA_FORMATTER.setUseCachedValuesForFormulaCells(true);
    }

    private final boolean skipEmptyRows;
    private final SpecialCharacterSanitizer specialCharSanitizer;

    protected ExcelSheetReader(boolean skipEmptyRows,
                               Set<CharSanitizeFlag> charSanitizeFlags) {
        this.skipEmptyRows = skipEmptyRows;
        this.specialCharSanitizer = new SpecialCharacterSanitizer(charSanitizeFlags);
    }

    /**
     * Create CSV data from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     *   each row will have the same number of columns
     */
    public String[][] convertToCsvData(Sheet sheet)  {
        List<String[]> excelListData = convertToCsvDataList(sheet);
        return excelListData.toArray(new String[0][0]);
    }

    public List<String[]> convertToCsvDataList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // NOTE: need to add 1 to the lastRowNum to make sure you don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(sheet, numOfRows);

        List<String[]> csvData = new ArrayList<>(numOfRows);

        // NOTE: avoid using "sheet.iterator" when looping through rows,
        //   b/c it can bail out early when it encounters the first empty line
        //   (even if there is more data rows remaining)
        for (int i = 0; i < numOfRows; i++) {
            String[] rowValues = new String[maxColumn];
            Row row = sheet.getRow(i);
            int columnCount = 0;
            // must check for null b/c a blank/empty row can (sometimes) return as null.
            if (row != null) {
                columnCount = Math.min( Math.max(row.getLastCellNum(), 0), maxColumn );
                for (int j = 0; j < columnCount; j++) {
                    rowValues[j] = getCellValue(row.getCell(j));
                }
            }

            // fill any 'extra' column cells with blank.
            for (int j = columnCount; j < maxColumn; j++) {
                rowValues[j] = "";
            }

            // ignore empty row if necessary
            if (this.skipEmptyRows && isEmptyRow(rowValues)) {
                continue;
            }
            csvData.add(rowValues);
        }
        return csvData;
    }

    /**
     * Iterate through the rows to find the max column width
     * @param sheet sheet
     * @param numOfRows total Number of Rows
     * @return max column
     */
    private int getMaxColumn(Sheet sheet, int numOfRows) {
        int maxColumn = 0;
        for (int i = 0; i < numOfRows; i++) {
            Row row = sheet.getRow(i);
            if (row != null) {
                int currentRowCellCount = row.getLastCellNum();
                if (currentRowCellCount > maxColumn) {
                    //  Sometimes a row is detected with more columns, but the 'extra'
                    //    column values are actually blank.  Therefore, double check if this is
                    //    the case and adjust accordingly.
                    for (int j = currentRowCellCount - 1; j >= maxColumn; j--) {
                        String cellValue = getCellValue(row.getCell(j));
                        if (! cellValue.isEmpty()) {
                            break;
                        }
                        currentRowCellCount--;
                    }
                    maxColumn = Math.max(maxColumn, currentRowCellCount);
                }
            }
        }
        return maxColumn;
    }

    private boolean isEmptyRow(String[] rowData) {
        for (String r : rowData) {
            if (r != null && !r.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     *  NOTE: dates & numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    private String getCellValue(Cell cell) {
        // note: the data formatter can handle a 'null' cell as well.
        String cellValue = EXCEL_DATA_FORMATTER.formatCellValue(cell);

        // if there are any certain special unicode characters (like nbsp or smart quotes),
        // replace w/ normal character equivalent
        cellValue = specialCharSanitizer.sanitize(cellValue);
        return cellValue.trim();
    }
}
