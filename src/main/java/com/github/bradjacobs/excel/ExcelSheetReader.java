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
import java.util.stream.Collectors;
import java.util.stream.IntStream;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag;

class ExcelSheetReader {
    protected final boolean skipEmptyRows;
    protected final CellValueReader cellValueReader;

    protected ExcelSheetReader(
            boolean autoTrim,
            boolean skipEmptyRows,
            Set<CharSanitizeFlag> charSanitizeFlags) {
        this.skipEmptyRows = skipEmptyRows;
        this.cellValueReader = new CellValueReader(autoTrim, charSanitizeFlags);
    }

    /**
     * Create 2-D data matrix from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     *   each row will have the same number of columns
     */
    public String[][] convertToMatrixData(Sheet sheet)  {
        List<String[]> excelListData = convertToMatrixDataList(sheet);
        return excelListData.toArray(new String[0][0]);
    }

    public List<String[]> convertToMatrixDataList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // grab all the rows from the sheet
        List<Row> rowList = getRows(sheet);
        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(rowList);
        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        return convertToMatrixDataList(rowList, availableColumns);
    }

    protected List<String[]> convertToMatrixDataList(List<Row> rowList, int[] availableColumns) {
        List<String[]> matrixDataList = new ArrayList<>(rowList.size());

        int totalColumnCount = availableColumns.length;
        int lastColumnIndex = totalColumnCount > 0 ? availableColumns[totalColumnCount-1] : -1;

        // NOTE: avoid using "sheet.iterator" when looping through rows,
        //   b/c it can bail out early when it encounters the first empty line
        //   (even if there is more data rows remaining)
        for (Row row : rowList) {
            String[] rowValues = new String[totalColumnCount];

            // must check for null because a blank/empty row can (sometimes) be null.
            int columnIndex = 0;
            if (row != null) {
                int rowColumnCount = Math.max(row.getLastCellNum(), 0);
                int lastRowColumnIndex = Math.min(rowColumnCount-1, lastColumnIndex);

                for (int availableColumn : availableColumns) {
                    if (availableColumn > lastRowColumnIndex) {
                        break;
                    }
                    rowValues[columnIndex++] = getCellValue(row.getCell(availableColumn));
                }
            }

            while (columnIndex < totalColumnCount) {
                rowValues[columnIndex++] = "";
            }

            // ignore empty row if necessary
            if (this.skipEmptyRows && isEmptyRow(rowValues)) {
                continue;
            }
            matrixDataList.add(rowValues);
        }

        // remove any trailing blank rows (even if 'skipEmptyRows==false')
        if (! this.skipEmptyRows) {
            while (!matrixDataList.isEmpty() && isEmptyRow(matrixDataList.get(matrixDataList.size()-1))) {
                matrixDataList.remove(matrixDataList.size()-1);
            }
        }

        return matrixDataList;
    }

    /**
     * Iterate through the rows to find the max column that contains a value
     * @param rowList list of rows for a sheet
     * @return max column
     */
    protected int getMaxColumn(List<Row> rowList) {
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

    protected boolean isEmptyRow(String[] rowData) {
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
    protected String getCellValue(Cell cell) {
        return cellValueReader.getCellValue(cell);
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *       (nulls are usually a 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return list of rows
     */
    protected List<Row> getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        return IntStream.range(0, numOfRows)
                .mapToObj(sheet::getRow)
                .collect(Collectors.toList());
    }

    /**
     * Determine which columns of the sheet should be read
     * @param sheet sheet
     * @param maxColumn max column count
     * @return int array of column indices to be read
     */
    protected int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        return IntStream.range(0, maxColumn).toArray();
    }
}
