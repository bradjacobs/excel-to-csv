/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.List;

import static com.github.bradjacobs.excel.util.RowDataUtil.toArray;
import static com.github.bradjacobs.excel.util.RowDataUtil.toUnmodifiableRow;
import static com.github.bradjacobs.excel.util.RowDataUtil.toUnmodifiableRows;


public class BasicSheetContent implements SheetContent {
    private static final String DEFAULT_SHEET_NAME = "";

    private final String sheetName;
    private final List<List<String>> rows;
    private final int rowCount;
    private final int columnCount;

    public BasicSheetContent(String[][] matrix) {
        this(DEFAULT_SHEET_NAME, matrix);
    }

    public BasicSheetContent(String sheetName, String[][] matrix) {
        this(sheetName, toUnmodifiableRows(matrix));
    }

    public BasicSheetContent(String sheetName, List<List<String>> rows) {
        this.sheetName = normalizeSheetName(sheetName);
        this.rows = toUnmodifiableRows(rows);
        this.rowCount = rows.size();
        this.columnCount = rowCount > 0 ? this.rows.get(0).size() : 0;
    }

    @Override
    public String getSheetName() {
        return sheetName;
    }

    @Override
    public int getRowCount() {
        return rowCount;
    }

    @Override
    public int getColumnCount() {
        return columnCount;
    }

    @Override
    public boolean isEmpty() {
        return rowCount == 0 || columnCount == 0;
    }

    @Override
    public String getCellValue(int rowIndex, int columnIndex) {
        validateRowIndex(rowIndex);
        validateColumnIndex(columnIndex);
        List<String> row = rows.get(rowIndex);
        return row.get(columnIndex);
    }

    @Override
    public List<String> getRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return toUnmodifiableRow(rows.get(rowIndex));
    }

    @Override
    public List<List<String>> getRows() {
        return toUnmodifiableRows(rows);
    }

    @Override
    public String[][] getMatrix() {
        return toArray(rows);
    }

    private static String normalizeSheetName(String sheetName) {
        return sheetName != null ? sheetName : DEFAULT_SHEET_NAME;
    }

    private void validateRowIndex(int rowIndex) {
        validateIndex(rowIndex, rowCount, "Row");
    }

    private void validateColumnIndex(int columnIndex) {
        validateIndex(columnIndex, columnCount, "Column");
    }

    private void validateIndex(int index, int size, String label) {
        if (index < 0 || index >= size) {
            throw new IndexOutOfBoundsException(
                    label + " index out of range: " + index + ", size: " + size
            );
        }
    }
}
