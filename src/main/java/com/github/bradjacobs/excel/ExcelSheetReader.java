/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;

class ExcelSheetReader {
    private final boolean skipEmptyRows;
    private final boolean skipInvisibleCells;
    private final CellValueReader cellValueReader;

    protected ExcelSheetReader(
            boolean autoTrim,
            boolean skipEmptyRows,
            boolean skipInvisibleCells,
            Set<CharSanitizeFlag> charSanitizeFlags) {
        this.skipEmptyRows = skipEmptyRows;
        this.skipInvisibleCells = skipInvisibleCells;
        this.cellValueReader = new CellValueReader(autoTrim, charSanitizeFlags);
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

        List<Row> rowList = getRows(sheet);

        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(rowList);
        int[] limitedColumns = getLimitedVisibleColumns(sheet, maxColumn);
        int totalColumnWidth = limitedColumns != null ? limitedColumns.length : maxColumn;

        List<String[]> csvData = new ArrayList<>(rowList.size());

        // NOTE: avoid using "sheet.iterator" when looping through rows,
        //   b/c it can bail out early when it encounters the first empty line
        //   (even if there is more data rows remaining)
        for (Row row : rowList) {
            String[] rowValues = new String[totalColumnWidth];

            // must check for null b/c a blank/empty row can (sometimes) be null.
            int j = 0;
            if (row != null) {
                // get the columnCount for the given row (can vary row by row)
                int columnCount = Math.min( Math.max(row.getLastCellNum(), 0), maxColumn );

                if (limitedColumns != null) {
                    for (int visibleColumn : limitedColumns) {
                        if (visibleColumn >= columnCount) {
                            break;
                        }
                        rowValues[j++] = getCellValue(row.getCell(visibleColumn));
                    }
                }
                else {
                    for (j = 0; j < columnCount; j++) {
                        rowValues[j] = getCellValue(row.getCell(j));
                    }
                }
            }

            while (j < totalColumnWidth) {
                rowValues[j++] = "";
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
     * Iterate through the rows to find the max column that contains a value
     * @param rowList list of rows for a sheet
     * @return max column
     */
    private int getMaxColumn(List<Row> rowList) {
        int maxColumn = 0;
        for (Row row : rowList) {
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
        return cellValueReader.getCellValue(cell);
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *       (nulls are usually a 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return list of rows
     */
    private List<Row> getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure you don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;
        List<Row> resultList = new ArrayList<>(numOfRows);

        for (int i = 0; i < numOfRows; i++) {
            Row row = sheet.getRow(i);
            if (this.skipInvisibleCells && row != null && row.getZeroHeight()) {
                continue;
            }
            resultList.add(row);
        }
        return resultList;
    }

    /**
     * Get any restricted column indexes to be considered for generating final data
     *   A 'null' is returned for no column restrictions
     * @param sheet Excel Sheet
     * @param maxColumn max columns
     * @return array of limited columns (or null for no limitation)
     */
    private int[] getLimitedVisibleColumns(Sheet sheet, int maxColumn) {
        if (this.skipInvisibleCells) {
            List<Integer> visibleColumnList = new ArrayList<>();
            for (int i = 0; i < maxColumn; i++) {
                boolean isColumnHidden = sheet.isColumnHidden(i);
                if (!isColumnHidden) {
                    visibleColumnList.add(i);
                }
            }

            if (visibleColumnList.size() < maxColumn) {
                return visibleColumnList.stream().mapToInt(Integer::intValue).toArray();
            }
        }

        // null means hidden columns are not considered (or none were actually found)
        return null;
    }
}
