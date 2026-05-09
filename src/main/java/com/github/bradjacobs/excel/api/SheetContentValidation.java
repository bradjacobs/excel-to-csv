package com.github.bradjacobs.excel.api;


final class SheetContentValidation {

    private SheetContentValidation() {
        // utility class
    }

    static void validateRectangularMatrix(String[][] matrix) {
        if (matrix == null || matrix.length == 0) {
            return;
        }

        int expectedColumnCount = getMatrixRowLength(matrix, 0);

        for (int rowIndex = 1; rowIndex < matrix.length; rowIndex++) {
            int actualColumnCount = getMatrixRowLength(matrix, rowIndex);

            if (actualColumnCount != expectedColumnCount) {
                throw new IllegalArgumentException(
                        "Matrix rows must all have the same number of columns. " +
                                "Expected " + expectedColumnCount +
                                " columns but found " + actualColumnCount +
                                " at row " + rowIndex + "."
                );
            }
        }
    }

    private static int getMatrixRowLength(String[][] matrix, int rowIndex) {
        String[] matrixRow = matrix[rowIndex];
        if (matrixRow == null) {
            throw new IllegalArgumentException("Matrix row must not be null at row " + rowIndex + ".");
        }
        return matrixRow.length;
    }
}