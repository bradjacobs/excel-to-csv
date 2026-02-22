/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.function.Consumer;
import java.util.function.UnaryOperator;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Reads an Excel Sheet and returns a 2-D array of data.
 */
public class ExcelSheetReader extends AbstractExcelSheetReader {
    protected final SheetConfig sheetConfig;
    protected final CellValueReader cellValueReader;

    // todo: still deciding if this constructor is ok or terrible.
    protected ExcelSheetReader(SheetConfig sheetConfig) {
        // override the internal POI utils size limit to allow for 'bigger Excel files'
        //   (as of POI version 5.2.0 the default value is 100_000_000)
        org.apache.poi.util.IOUtils.setByteArrayMaxOverride(Integer.MAX_VALUE);

        this.sheetConfig = sheetConfig;
        this.cellValueReader = new CellValueReader(sheetConfig.isAutoTrim(), sheetConfig.getCharSanitizeFlags());
    }

    // interface for grabbing a specific sheet from the workbook.
    @FunctionalInterface
    private interface WorkbookSheetGrabber {
        Sheet extractSheet(Workbook workbook);
    }

    @Override
    public String[][] readExcelSheetData(InputStream is, int sheetIndex, String password) throws IOException {
        // check negative before trying to read the file.
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("SheetIndex cannot be negative");
        }
        Sheet sheet = getFileSheet(is, password, (w) -> w.getSheetAt(sheetIndex));
        return convertToDataMatrix(sheet);
    }

    @Override
    public String[][] readExcelSheetData(InputStream is, String sheetName, String password) throws IOException {
        // Note: passing in a 'null' sheetName can cause NPE, but will be fixed in the next POI release.
        //   https://github.com/apache/poi/commit/04f4c1fa7424f12b12f1e513950f9e7fa13c625d
        Sheet sheet = getFileSheet(is, password, (w) -> w.getSheet(sheetName));
        if (sheet == null) {
            throw new IllegalArgumentException(String.format("Unable to find sheet with name: '%s'", sheetName));
        }
        return convertToDataMatrix(sheet);
    }

    /**
     * Read in a specific Sheet from the Excel File input stream.
     * @param inputStream inputStream of the excel file.
     * @param password password (optional)
     * @param workbookSheetGrabber function to grab the sheet from the workbook.
     * @return sheet
     */
    private Sheet getFileSheet(InputStream inputStream,
                               String password,
                               WorkbookSheetGrabber workbookSheetGrabber
    ) throws IOException {
        try (inputStream; Workbook wb = WorkbookFactory.create(inputStream, password)) {
            return workbookSheetGrabber.extractSheet(wb);
        }
    }

    /**
     * Create 2-D data matrix from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     * each row will have the same number of columns
     */
    public String[][] convertToDataMatrix(Sheet sheet) {
        List<String[]> excelListData = convertToDataMatrixList(sheet);
        return excelListData.toArray(new String[0][0]);
    }

    public List<String[]> convertToDataMatrixList(Sheet sheet) {
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // grab all the rows from the sheet
        List<Row> rowList = getRows(sheet);
        // first scan the rows to find the max column width
        int maxColumn = getMaxColumn(sheet, rowList);
        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        return convertToDataMatrixList(rowList, availableColumns);
    }

    protected List<String[]> convertToDataMatrixList(List<Row> rowList, int[] columnsToRead) {
        // if there are no available columns then bail early.
        if (columnsToRead.length == 0) {
            return Collections.emptyList();
        }

        final int columnCount = columnsToRead.length;
        final int maxRequestedColumnIndex = columnsToRead[columnCount-1];

        StringArrayRowConsumer stringArrayRowConsumer =
                new StringArrayRowConsumer(sheetConfig.isRemoveBlankRows(), sheetConfig.isRemoveBlankColumns());
        for (Row row : rowList) {
            String[] rowValues = toRowValues(row, columnsToRead, columnCount, maxRequestedColumnIndex);
            stringArrayRowConsumer.accept(rowValues);
        }

        // call finalizeRows to handle any extra remove blank row/column logic
        stringArrayRowConsumer.finalizeRows();

        // lastly return all the rows
        return stringArrayRowConsumer.getRows();
    }

    /**
     * Converts the given row into String[] of values.
     * @param row excel sheet row
     * @param columnsToRead columns to read
     * @param columnCount column count
     * @param maxRequestedColumnIndex max requested column index
     * @return String[] of values
     */
    private String[] toRowValues(Row row,
                                 int[] columnsToRead,
                                 int columnCount,
                                 int maxRequestedColumnIndex) {
        String[] rowValues = new String[columnCount];
        int columnIndex = 0;

        // must check for null because a blank/empty row can (sometimes) be null.
        if (row != null) {
            int rowCellCount = Math.max(row.getLastCellNum(), 0);
            int lastRowColumnIndex = Math.min(rowCellCount-1, maxRequestedColumnIndex);

            for (int currentColumnToRead : columnsToRead) {
                if (currentColumnToRead > lastRowColumnIndex) {
                    break;
                }
                rowValues[columnIndex++] = getCellValue(row.getCell(currentColumnToRead));
            }
        }

        if (columnIndex < columnCount) {
            Arrays.fill(rowValues, columnIndex, columnCount, "");
        }
        return rowValues;
    }

    /**
     * Iterate through the rows to find the max column that contains a value
     * @param rowList list of rows for a sheet
     * @return max column
     */
    protected int getMaxColumn(Sheet sheet, List<Row> rowList) {
        int maxColumnCount = 0;
        for (Row row : rowList) {
            if (row == null) {
                continue;
            }

            int currentRowCellCount = row.getLastCellNum();
            if (currentRowCellCount <= maxColumnCount) {
                continue;
            }

            //  Sometimes a row is detected with more columns, but the 'extra'
            //    column values are actually blank.  Therefore, double check if this is
            //    the case and adjust accordingly.
            for (int j = currentRowCellCount - 1; j >= maxColumnCount; j--) {
                String cellValue = getCellValue(row.getCell(j));
                if (!cellValue.isEmpty()) {
                    // if a new max column candidate has a value,
                    // but the cell is invisible and configured to
                    // ignore invisible cells, then should continue on.
                    // todo: look for less kludgy way to solve this
                    if (!sheetConfig.isRemoveInvisibleCells() || !sheet.isColumnHidden(j)) {
                        break;
                    }
                }
                currentRowCellCount--;
            }
            maxColumnCount = Math.max(maxColumnCount, currentRowCellCount);
        }
        return maxColumnCount;
    }

    /**
     * Gets the string representation of the value in the cell
     *   (where the cell value is what you "physically see" in the cell)
     * NOTE: dates & numbers should retain their original formatting.
     * @param cell excel cell
     * @return string representation of the cell.
     */
    protected String getCellValue(Cell cell) {
        return cellValueReader.getCellValue(cell);
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *   (nulls are usually 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return list of rows
     */
    protected List<Row> getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int numOfRows = sheet.getLastRowNum() + 1;

        return IntStream.range(0, numOfRows)
                .mapToObj(sheet::getRow)
                .filter(this::isRowVisible)
                .collect(Collectors.toList());
    }

    protected boolean isRowVisible(Row row) {
        if (!sheetConfig.isRemoveInvisibleCells()) {
            return true;
        }
        return row == null || !row.getZeroHeight();
    }

    /**
     * Determine which columns of the sheet should be read
     * @param sheet sheet
     * @param maxColumn max column count
     * @return int array of column indices to be read
     */
    protected int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        if (!sheetConfig.isRemoveInvisibleCells()) {
            return IntStream.range(0, maxColumn).toArray();
        }

        return IntStream.range(0, maxColumn)
                .filter(columnIndex -> !sheet.isColumnHidden(columnIndex))
                .toArray();
    }

    /**
     * Consumer class to add the String[] rows, which will handle
     * any additional logic regarding the removal of blank rows or columns (as needed)
     */
    private static class StringArrayRowConsumer implements Consumer<String[]> {
        private final List<String[]> rowList = new ArrayList<>();
        private final boolean removeBlankRows; // flag to remove blank rows
        private final boolean removeBlankColumns; // flag to remove blank columns
        private boolean[] columnsWithDataFlags = null; // tracks which columns have a value
        private int columnsWithDataCount = 0; // tracks total number of columns that have non-blank value.

        public StringArrayRowConsumer(boolean removeBlankRows, boolean removeBlankColumns) {
            this.removeBlankRows = removeBlankRows;
            this.removeBlankColumns = removeBlankColumns;
        }

        /**
         * Consumes all the String[] rows read from the Excel file
         * _ASSERT_ that all rows passed in will have the same length.
         * @param row String[] row
         */
        @Override
        public void accept(String[] row) {
            // Remove blank rows if configured
            if (this.removeBlankRows && isEmptyRow(row))
                return;

            // if removing blank columns, then keep track which columns contain values.
            if (this.removeBlankColumns) {
                if (columnsWithDataFlags == null) {
                    columnsWithDataFlags = new boolean[row.length];
                }
                if (columnsWithDataCount < columnsWithDataFlags.length) {
                    for (int i = 0; i < columnsWithDataFlags.length; i++) {
                        if (!columnsWithDataFlags[i] && !row[i].isEmpty()) {
                            columnsWithDataFlags[i] = true;
                            columnsWithDataCount++;
                        }
                    }
                }
            }
            this.rowList.add(row);
        }

        public List<String[]> getRows() {
            return Collections.unmodifiableList(rowList);
        }

        /**
         * To be called after all rows have been added for any post-processing (if necessary)
         */
        public void finalizeRows() {
            // remove any trailing blank rows (regardless of 'removeBlankRows' value)
            while (!rowList.isEmpty() && isEmptyRow(rowList.get(rowList.size()-1))) {
                rowList.remove(rowList.size()-1);
            }

            // remove any blank columns (if necessary)
            if (removeBlankColumns &&
                    !rowList.isEmpty() &&
                    columnsWithDataFlags != null &&
                    columnsWithDataCount < columnsWithDataFlags.length) {
                rowList.replaceAll(new FilterColumnsOperator(columnsWithDataFlags, columnsWithDataCount));
            }
        }
    }

    // Helper class to remove blank column(s) from each String[] row
    private static class FilterColumnsOperator implements UnaryOperator<String[]> {
        private final boolean[] columnsWithDataFlags;
        private final int columnsWithDataCount;

        public FilterColumnsOperator(boolean[] columnsWithDataFlags, int columnsWithDataCount) {
            this.columnsWithDataFlags = columnsWithDataFlags;
            this.columnsWithDataCount = columnsWithDataCount;
        }

        @Override
        public String[] apply(String[] row) {
            String[] filtered = new String[columnsWithDataCount];
            for (int i = 0, idx = 0; i < columnsWithDataFlags.length && i < row.length; i++) {
                if (columnsWithDataFlags[i]) {
                    filtered[idx++] = row[i];
                }
            }
            return filtered;
        }
    }

    /**
     * Check if row is considered to be 'blank/empty'
     * @param rowData data row
     * @return true if 'empty'
     */
    protected static boolean isEmptyRow(String[] rowData) {
        for (String r : rowData) {
            if (r != null && !r.isEmpty()) {
                return false;
            }
        }
        return true;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractSheetConfigBuilder<ExcelSheetReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public ExcelSheetReader build() {
            return new ExcelSheetReader(this.buildConfig());
        }
    }
}
