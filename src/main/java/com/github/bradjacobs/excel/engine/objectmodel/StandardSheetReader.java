/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.objectmodel;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.engine.rows.StringRowConsumer;
import com.github.bradjacobs.excel.model.BasicSheetContent;
import com.github.bradjacobs.excel.model.SheetContent;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

public class StandardSheetReader {
    private final SheetConfig sheetConfig;
    private final CellValueReader cellValueReader;

    // TODO - consider alternate ways to construct this
    public StandardSheetReader(SheetConfig sheetConfig) {
        this.sheetConfig = sheetConfig;
        this.cellValueReader = new CellValueReader(sheetConfig.trimStringValues(), sheetConfig.getCharSanitizeFlags());
    }

    public SheetContent toSheetContent(Sheet sheet) {
        Validate.isTrue(sheet != null, "Sheet parameter cannot be null.");

        // grab all the rows from the sheet
        RowInfo rowInfo = getRows(sheet);
        List<Row> rowList = rowInfo.rowList;
        int maxColumn = rowInfo.maxColumn;

        // get all the column (indexes) that are to be read
        int[] availableColumns = getAvailableColumns(sheet, maxColumn);

        List<List<String>> sheetDataRows = convertToSheetDataRows(rowList, availableColumns);
        return new BasicSheetContent(sheet.getSheetName(), sheetDataRows);
    }

    private List<List<String>> convertToSheetDataRows(List<Row> rowList, int[] columnsToRead) {
        // if there are no available columns then bail early.
        if (columnsToRead.length == 0) {
            return List.of();
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

        return stringRowConsumer.generateRowDataList();
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
     * Determine which columns of the sheet should be read
     * @param sheet sheet
     * @param maxColumn max column count
     * @return int array of column indices to be read
     */
    private int[] getAvailableColumns(Sheet sheet, int maxColumn) {
        return IntStream.range(0, maxColumn)
                .filter(columnIndex -> shouldIncludeColumn(sheet, columnIndex))
                .toArray();
    }

    /**
     * Method to grab all the rows for the sheet ahead of time
     *   NOTE: some elements in the result list could be 'null'
     *   (nulls are usually 'default unaltered rows')
     * @param sheet input Excel Sheet
     * @return rowInfo containing a list of rows and the max column detected
     */
    private RowInfo getRows(Sheet sheet) {
        // NOTE: need to add 1 to the lastRowNum to make sure don't skip the last row
        //  (however doesn't seem to need for this when using row.getLastCellNum, which seems odd)
        int rowCount = sheet.getLastRowNum() + 1;

        // Note: avoid using 'sheet.iterator()', as it can
        //   have problems when encountering 'null' rows.
        int maxColumnCount = 0;
        List<Row> rowList = new ArrayList<>(rowCount);
        for (int i = 0; i < rowCount; i++) {
            Row row = sheet.getRow(i);
            if (shouldIncludeRow(row)) {
                rowList.add(row);
                maxColumnCount = Math.max(maxColumnCount, getColumnCount(row));
            }
        }
        return new RowInfo(rowList, maxColumnCount);
    }

    private int getColumnCount(Row row) {
        int rowLastCellNumber = row != null ? row.getLastCellNum() : 0;
        return Math.max(rowLastCellNumber, 0);
    }


    /**
     * Determines whether a row should be included in the generated sheet content.
     * <p>
     * When hidden cells are not being skipped, all rows are included.
     * When hidden cells are being skipped, hidden rows are excluded.
     *
     * @param row Excel row to evaluate; may be {@code null} for default, unaltered rows
     * @return {@code true} if the row should be included; {@code false} otherwise
     */
    private boolean shouldIncludeRow(Row row) {
        return !(sheetConfig.skipHiddenCells() && isHiddenRow(row));
    }

    private boolean isHiddenRow(Row row) {
        return row != null && row.getZeroHeight();
    }

    private boolean shouldIncludeColumn(Sheet sheet, int columnIndex) {
        return !(sheetConfig.skipHiddenCells() && isHiddenColumn(sheet, columnIndex));
    }

    private static boolean isHiddenColumn(Sheet sheet, int columnIndex) {
        return sheet.isColumnHidden(columnIndex);
    }

    /**
     * Simple POJO to hold Excel sheet rows and size of largest row.
     */
    private static class RowInfo {
        private final List<Row> rowList;
        private final int maxColumn;

        public RowInfo(List<Row> rowList, int maxColumn) {
            this.rowList = rowList;
            this.maxColumn = maxColumn;
        }
    }
}
