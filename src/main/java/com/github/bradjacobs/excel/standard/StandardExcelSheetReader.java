/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.api.BasicSheetContent;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader;
import com.github.bradjacobs.excel.core.CellValueReader;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.request.SheetInfo;
import com.github.bradjacobs.excel.request.SheetSelector;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Collectors;
import java.util.stream.IntStream;

/**
 * Reads an Excel Sheet and returns a 2-D array of data.
 */
public class StandardExcelSheetReader extends AbstractExcelSheetReader {
    private final CellValueReader cellValueReader;

    // todo: still deciding if this constructor is ok or terrible.
    public StandardExcelSheetReader(SheetConfig sheetConfig) {
        super(sheetConfig);
        this.cellValueReader = new CellValueReader(sheetConfig.trimStringValues(), sheetConfig.getCharSanitizeFlags());
    }

    @Override
    public List<SheetContent> readSheets(ExcelSheetReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");

        String password = request.getPassword();
        SheetSelector sheetSelector = request.getSheetSelector();
        InputStream excelInputStream = request.getSourceInputStream();

        try (excelInputStream; Workbook workbook = WorkbookFactory.create(excelInputStream, password)) {
            List<WorkbookSheetInfo> selectedSheets = selectWorkbookSheets(workbook, sheetSelector);
            return toSheetContents(selectedSheets);
        }
    }

    private List<WorkbookSheetInfo> selectWorkbookSheets(Workbook workbook, SheetSelector sheetSelector) {
        List<WorkbookSheetInfo> workbookSheets = readWorkbookSheets(workbook);
        return sheetSelector.filterSheets(workbookSheets);
    }

    private List<SheetContent> toSheetContents(List<WorkbookSheetInfo> selectedSheets) {
        return selectedSheets.stream()
                .map(this::toSheetContent)
                .collect(Collectors.toList());
    }

    /**
     * Gets all sheets in the given workbook
     * returns a list of 'WorkbookSheetInfo', which includes: Sheet, SheetName, SheetIndex.
     * @param workbook workbook
     * @return list of workbook sheets
     */
    private List<WorkbookSheetInfo> readWorkbookSheets(Workbook workbook) {
        int sheetCount = workbook.getNumberOfSheets();
        List<WorkbookSheetInfo> sheetInfos = new ArrayList<>(sheetCount);

        for (int sheetIndex = 0; sheetIndex < sheetCount; sheetIndex++) {
            Sheet sheet = workbook.getSheetAt(sheetIndex);
            sheetInfos.add(new WorkbookSheetInfo(sheet.getSheetName(), sheetIndex, sheet));
        }
        return sheetInfos;
    }

    private SheetContent toSheetContent(WorkbookSheetInfo sheetInfo) {
        String[][] sheetDataMatrix = convertToSheetContentArray(sheetInfo.getSheet());
        return new BasicSheetContent(sheetInfo.getName(), sheetDataMatrix);
    }

    // todo: change to non-public unless reason to keep public.
    /**
     * Create 2-D data matrix from the given Excel Sheet
     * @param sheet Excel Sheet
     * @return 2-D array representing CSV format
     * each row will have the same number of columns
     */
    public String[][] convertToSheetContentArray(Sheet sheet) {
        Validate.isTrue(sheet != null, "Sheet parameter cannot be null.");

        // grab all the rows from the sheet
        RowInfo rowInfo = getRows(sheet);
        List<Row> rowList = rowInfo.rowList;
        int maxColumn = rowInfo.maxColumn;

        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        return convertToSheetContentArray(rowList, availableColumns);
    }

    private String[][] convertToSheetContentArray(List<Row> rowList, int[] columnsToRead) {
        // if there are no available columns then bail early.
        if (columnsToRead.length == 0) {
            return new String[0][0];
        }

        final int columnCount = columnsToRead.length;
        final int maxRequestedColumnIndex = columnsToRead[columnCount-1];

        StringRowConsumer stringRowConsumer =
                StringRowConsumer.of(
                        sheetConfig.skipBlankRows(),
                        sheetConfig.skipBlankColumns());

        for (Row row : rowList) {
            List<String> rowValuesList = toRowValues(row, columnsToRead, maxRequestedColumnIndex);
            stringRowConsumer.accept(rowValuesList);
        }

        return stringRowConsumer.generateMatrix();
    }

    /**
     * Converts the given row into String[] of values.
     * @param row Excel sheet row
     * @param columnsToRead columns to read
     * @param maxRequestedColumnIndex max requested column index
     * @return List of values
     */
    private List<String> toRowValues(Row row,
                                 int[] columnsToRead,
                                 int maxRequestedColumnIndex) {
        List<String> rowValues = new ArrayList<>();

        int rowCellCount = getColumnCount(row);
        if (rowCellCount > 0) {
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
    private String getCellValue(Cell cell) {
        return cellValueReader.getCellValue(cell);
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *   (nulls are usually 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return list of rows and the max column detected
     */
    private RowInfo getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int rowCount = sheet.getLastRowNum() + 1;

        int maxColumnCount = 0;
        List<Row> rowList = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (isRowVisible(row)) {
                rowList.add(row);
                maxColumnCount = Math.max(maxColumnCount, getColumnCount(row));
            }
        }
        return new RowInfo(rowList, maxColumnCount);
    }

    private int getColumnCount(Row row) {
        if (row == null) {
            return 0;
        }
        short lastCellIndexExclusive = row.getLastCellNum();
        return Math.max(lastCellIndexExclusive, 0);
    }


    private boolean isRowVisible(Row row) {
        if (!sheetConfig.skipInvisibleCells()) {
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
    private int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        if (!sheetConfig.skipInvisibleCells()) {
            return IntStream.range(0, maxColumn).toArray();
        }

        return IntStream.range(0, maxColumn)
                .filter(columnIndex -> !sheet.isColumnHidden(columnIndex))
                .toArray();
    }

    private static class RowInfo {
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

    public static class Builder extends AbstractSheetConfigBuilder<StandardExcelSheetReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public StandardExcelSheetReader build() {
            return new StandardExcelSheetReader(this.buildConfig());
        }
    }


    private static class WorkbookSheetInfo implements SheetInfo {
        private final String sheetName;
        private final int sheetIndex;
        private final Sheet sheet;

        public WorkbookSheetInfo(String sheetName, int sheetIndex, Sheet sheet) {
            this.sheetName = sheetName;
            this.sheetIndex = sheetIndex;
            this.sheet = sheet;
        }

        @Override
        public String getName() {
            return sheetName;
        }

        @Override
        public int getIndex() {
            return sheetIndex;
        }

        public Sheet getSheet() {
            return sheet;
        }
    }
}
