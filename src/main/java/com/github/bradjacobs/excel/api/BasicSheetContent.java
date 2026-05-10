/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.util.Arrays;
import java.util.Collections;
import java.util.List;
import java.util.stream.Collectors;


public class BasicSheetContent implements SheetContent {
    private static final String DEFAULT_SHEET_NAME = "";
    private static final String[][] EMPTY_MATRIX = new String[0][0];

    private final String sheetName;
    private final String[][] matrix;
    private final int rowCount;
    private final int columnCount;

    public static SheetContent fromMatrix(String sheetName, String[][] matrix) {
        return new BasicSheetContent(sheetName, matrix);
    }

    public BasicSheetContent(String[][] matrix) {
        this(DEFAULT_SHEET_NAME, matrix);
    }

    public BasicSheetContent(String sheetName, String[][] matrix) {
        this.sheetName = normalizeSheetName(sheetName);
        this.matrix = normalizeMatrix(matrix);
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
        return toUnmodifiableRow(matrix[rowIndex]);
    }

    @Override
    public List<List<String>> getRows() {
        return Arrays.stream(matrix)
                .map(this::toUnmodifiableRow)
                .collect(Collectors.collectingAndThen(
                        Collectors.toList(),
                        Collections::unmodifiableList));
    }

    @Override
    public String[][] getMatrix() {

        // TODO - return a copy of the matrix
        //   mainly to avoid a 'spotbugs' warning
        //   but need an alternate solution to avoid
        //   copying the data so much
        String[][] copy = new String[matrix.length][];

        for (int i = 0; i < matrix.length; i++) {
            String[] row = matrix[i];
            copy[i] = Arrays.copyOf(row, row.length);
        }

        return copy;
    }

    private static String normalizeSheetName(String sheetName) {
        return sheetName != null ? sheetName : DEFAULT_SHEET_NAME;
    }

    private static String[][] normalizeMatrix(String[][] matrix) {
        SheetContentValidation.validateRectangularMatrix(matrix);
        return matrix != null ? matrix : EMPTY_MATRIX;
    }

    private List<String> toUnmodifiableRow(String[] row) {
        return Collections.unmodifiableList(Arrays.asList(row));
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
