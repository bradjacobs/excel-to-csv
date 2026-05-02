/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;


public class BasicSheetContent implements SheetContent  {
    private final String sheetName;
    private final String[][] matrix;
    private final int rowCount;
    private final int columnCount;

    public static SheetContent fromMatrix(String sheetName, String[][] matrix) {
        return new BasicSheetContent(sheetName, matrix);
    }

    public BasicSheetContent(String[][] matrix) {
        this("", matrix);
    }

    public BasicSheetContent(String sheetName, String[][] matrix) {
        this.sheetName = sheetName != null ? sheetName : "";
        this.matrix = matrix != null ? matrix : new String[0][0];
        this.rowCount = this.matrix.length;
        this.columnCount = rowCount > 0 ? this.matrix[0].length : 0;
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
        return matrix[rowIndex][columnIndex];
    }

    @Override
    public List<String> getRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return toReadOnlyRow(matrix[rowIndex]);
    }

    @Override
    public List<List<String>> getRows() {
        return Arrays.stream(matrix)
                .map(this::toReadOnlyRow)
                .collect(Collectors.collectingAndThen(
                        Collectors.toList(),
                        Collections::unmodifiableList));
    }

    private List<String> toReadOnlyRow(String[] row) {
        List<String> readOnlyRow = Arrays.asList(row);
        return Collections.unmodifiableList(readOnlyRow);
    }
    
    @Override
    public String[][] getMatrix() {
        // NOTE: going to forgo making a full copy
        // just for readonly-protection
        return matrix;
    }

    private void validateRowIndex(int rowIndex) {
        if (rowIndex < 0 || rowIndex >= rowCount) {
            throw new IndexOutOfBoundsException(
                    "Row index out of range: " + rowIndex + ", size: " + rowCount
            );
        }
    }

    private void validateColumnIndex(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= columnCount) {
            throw new IndexOutOfBoundsException(
                    "Column index out of range: " + columnIndex + ", size: " + columnCount
            );
        }
    }
}
