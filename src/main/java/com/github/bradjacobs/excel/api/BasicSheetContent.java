/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;


public class BasicSheetContent implements SheetContent {
    private final String sheetName;
    private final String[][] dataMatrix;
    private final int rowCount;
    private final int columnCount;

    public static BasicSheetContent of(String sheetName, String[][] dataMatrix) {
        return new BasicSheetContent(sheetName, dataMatrix);
    }

    public BasicSheetContent(String[][] dataMatrix) {
        this("", dataMatrix);
    }

    public BasicSheetContent(String sheetName, String[][] dataMatrix) {
        this.sheetName = sheetName != null ? sheetName : "";
        this.dataMatrix = dataMatrix != null ? dataMatrix : new String[0][0];
        this.rowCount = this.dataMatrix.length;
        this.columnCount = rowCount > 0 ? this.dataMatrix[0].length : 0;
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
        return dataMatrix[rowIndex][columnIndex];
    }

    @Override
    public List<String> getRow(int rowIndex) {
        validateRowIndex(rowIndex);
        return List.copyOf(Arrays.asList(dataMatrix[rowIndex]));
    }

    @Override
    public List<List<String>> getRows() {
        return Arrays.stream(dataMatrix)
                .map(Arrays::asList)
                .collect(Collectors.toList());
    }

    @Override
    public String[][] getMatrix() {
        // NOTE: going to forgo making a full copy
        // just for readonly-protection
        return dataMatrix;
    }

    // todo better message
    private void validateRowIndex(int rowIndex) {
        if (rowIndex < 0 || rowIndex >= rowCount) {
            throw new IndexOutOfBoundsException(
                    "Row index out of bounds: " + rowIndex + " (size: " + rowCount + ")"
            );
        }
    }

    private void validateColumnIndex(int columnIndex) {
        if (columnIndex < 0 || columnIndex >= columnCount) {
            throw new IndexOutOfBoundsException(
                    "Column index out of bounds: " + columnIndex + " (size: " + columnCount + ")"
            );
        }
    }
}
