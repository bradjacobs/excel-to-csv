/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * Verifies the matrix shape contract enforced before sheet content is accepted.
 */
class SheetContentValidationTest {

    private static final String NULL_ROW_MESSAGE_TEMPLATE = "Matrix row must not be null at row %d.";
    private static final String COLUMN_COUNT_MISMATCH_MESSAGE_TEMPLATE =
            "Matrix rows must all have the same number of columns. Expected %d columns but found %d at row %d.";

    @ParameterizedTest(name = "{0}")
    @MethodSource("validMatrixScenarios")
    void validateRectangularMatrixAllowsValidMatrices(String displayName, String[][] matrix) {
        assertValidMatrix(matrix);
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("invalidMatrixScenarios")
    void validateRectangularMatrixThrowsForInvalidMatrices(
            String displayName,
            String[][] matrix,
            String expectedMessage
    ) {
        assertInvalidMatrix(matrix, expectedMessage);
    }

    /**
     * Supplies matrices that are either rectangular or intentionally accepted as empty input.
     */
    private static Stream<Arguments> validMatrixScenarios() {
        return Stream.of(
                Arguments.of("null matrix", null),
                Arguments.of("empty matrix", new String[][]{}),
                Arguments.of("single empty row", new String[][]{
                        {}
                }),
                Arguments.of("single populated row", new String[][]{
                        {"A1", "B1", "C1"}
                }),
                Arguments.of("rectangular matrix", new String[][]{
                        {"A1", "B1", "C1"},
                        {"A2", "B2", "C2"},
                        {"A3", "B3", "C3"}
                }),
                Arguments.of("all rows empty", new String[][]{
                        {},
                        {},
                        {}
                }),
                Arguments.of("null cell values", new String[][]{
                        {"A1", null, "C1"},
                        {null, "B2", "C2"},
                        {"A3", "B3", null}
                })
        );
    }

    /**
     * Supplies malformed matrices with the exact validation message expected for each failure.
     */
    private static Stream<Arguments> invalidMatrixScenarios() {
        return Stream.of(
                Arguments.of(
                        "first row is null",
                        new String[][]{
                                null,
                                {"A2", "B2"}
                        },
                        expectedNullRowMessage(0)
                ),
                Arguments.of(
                        "later row is null",
                        new String[][]{
                                {"A1", "B1"},
                                null,
                                {"A3", "B3"}
                        },
                        expectedNullRowMessage(1)
                ),
                Arguments.of(
                        "later row has fewer columns",
                        new String[][]{
                                {"A1", "B1", "C1"},
                                {"A2", "B2"}
                        },
                        expectedColumnCountMismatchMessage(3, 2, 1)
                ),
                Arguments.of(
                        "later row has more columns",
                        new String[][]{
                                {"A1", "B1"},
                                {"A2", "B2", "C2"}
                        },
                        expectedColumnCountMismatchMessage(2, 3, 1)
                ),
                Arguments.of(
                        "first row is empty and later row has columns",
                        new String[][]{
                                {},
                                {"A2"}
                        },
                        expectedColumnCountMismatchMessage(0, 1, 1)
                ),
                Arguments.of(
                        "throws at first non-rectangular row",
                        new String[][]{
                                {"A1", "B1"},
                                {"A2", "B2"},
                                {"A3"},
                                {"A4", "B4", "C4"}
                        },
                        expectedColumnCountMismatchMessage(2, 1, 2)
                )
        );
    }

    private static void assertValidMatrix(String[][] matrix) {
        assertDoesNotThrow(() -> SheetContentValidation.validateRectangularMatrix(matrix));
    }

    private static void assertInvalidMatrix(String[][] matrix, String expectedMessage) {
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            SheetContentValidation.validateRectangularMatrix(matrix);
        });
        assertEquals(expectedMessage, exception.getMessage());
    }

    private static String expectedNullRowMessage(int rowIndex) {
        return String.format(NULL_ROW_MESSAGE_TEMPLATE, rowIndex);
    }

    private static String expectedColumnCountMismatchMessage(int expectedColumnCount, int actualColumnCount, int rowIndex) {
        return String.format(
                COLUMN_COUNT_MISMATCH_MESSAGE_TEMPLATE,
                expectedColumnCount,
                actualColumnCount,
                rowIndex
        );
    }
}
