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
import java.util.List;
import java.util.stream.IntStream;

/**
 * Reads an Excel Sheet and returns a 2-D array of data.
 */
public class ExcelSheetReader extends AbstractExcelSheetReader {
    protected final SheetConfig sheetConfig;
    protected final CellValueReader cellValueReader;

    // todo: still deciding if this constructor is ok or terrible.
    protected ExcelSheetReader(SheetConfig sheetConfig) {
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
        if (sheet == null) {
            throw new IllegalArgumentException("Sheet parameter cannot be null.");
        }

        // grab all the rows from the sheet
        RowInfo rowInfo = getRows(sheet);
        List<Row> rowList = rowInfo.rowList;
        int maxColumn = rowInfo.maxColumn;

        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        return convertToDataMatrix(rowList, availableColumns);
    }

    protected String[][] convertToDataMatrix(List<Row> rowList, int[] columnsToRead) {
        // if there are no available columns then bail early.
        if (columnsToRead.length == 0) {
            return new String[0][0];
        }

        final int columnCount = columnsToRead.length;
        final int maxRequestedColumnIndex = columnsToRead[columnCount-1];

        StringRowConsumer stringRowConsumer = StringRowConsumer.of(sheetConfig.isRemoveBlankRows(), sheetConfig.isRemoveBlankColumns());

        for (Row row : rowList) {
            List<String> rowValuesList = toRowValues(row, columnsToRead, maxRequestedColumnIndex);
            stringRowConsumer.accept(rowValuesList);
        }

        return stringRowConsumer.generateMatrix();
    }

    /**
     * Converts the given row into String[] of values.
     * @param row excel sheet row
     * @param columnsToRead columns to read
     * @param maxRequestedColumnIndex max requested column index
     * @return List of values
     */
    private List<String> toRowValues(Row row,
                                 int[] columnsToRead,
                                 int maxRequestedColumnIndex) {
        List<String> rowValues = new ArrayList<>();

        // must check for null because a blank/empty row can (sometimes) be null.
        if (row != null) {
            int rowCellCount = Math.max(row.getLastCellNum(), 0);
            int lastRowColumnIndex = Math.min(rowCellCount-1, maxRequestedColumnIndex);

            for (int currentColumnToRead : columnsToRead) {
                if (currentColumnToRead > lastRowColumnIndex) {
                    break;
                }
                rowValues.add(getCellValue(row.getCell(currentColumnToRead)));
            }
        }

        return rowValues;
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
     * @return list of rows and the max column detected
     */
    protected RowInfo getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int rowCount = sheet.getLastRowNum() + 1;

        int maxColumnCount = 0;
        List<Row> rowList = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (isRowVisible(row)) {
                rowList.add(row);
                maxColumnCount = Math.max(maxColumnCount, row != null ? row.getLastCellNum() : 0);
            }
        }
        return new RowInfo(rowList, maxColumnCount);
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

    protected static class RowInfo {
        protected final List<Row> rowList;
        protected final int maxColumn;

        public RowInfo(List<Row> rowList, int maxColumn) {
            this.rowList = rowList;
            this.maxColumn = maxColumn;
        }
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
