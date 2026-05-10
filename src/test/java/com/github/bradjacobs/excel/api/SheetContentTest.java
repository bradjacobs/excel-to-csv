/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
abstract class SheetContentTest {

    protected static final String BLANK = "";
    protected static final String[][] INPUT_MATRIX = {
            {"dog", "cat", "bird"},
            {"frog", "cow", "elephant"}
    };
    protected static final List<List<String>> INPUT_LIST = Arrays.stream(INPUT_MATRIX)
            .map(Arrays::asList)
            .collect(Collectors.toList());
    protected static final String INPUT_SHEET_NAME = "mySheet";

    protected static List<List<String>> copyOfInput() {
        return copyOfInput(null);
    }

    protected static List<List<String>> copyOfInput(Integer minWidth) {
        List<List<String>> copy = new ArrayList<>();
        for (List<String> inner : INPUT_LIST) {
            List<String> copyInnerRow = new ArrayList<>(inner);
            if (minWidth != null) {
                while (copyInnerRow.size() < minWidth) {
                    copyInnerRow.add(BLANK);
                }
            }
            copy.add(copyInnerRow);
        }
        return copy;
    }

    abstract protected SheetContent createDefaultSheetContent();
    abstract protected SheetContent createEmptySheetContent();

    @Nested
    @DisplayName("Default Get Behavior")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class DefaultGetTests {

        @Test
        public void getSheetName() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_SHEET_NAME, sheetContent.getSheetName());
        }

        @Test
        public void getRowCount() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_MATRIX.length, sheetContent.getRowCount());
        }

        @Test
        public void getColumnCount() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_MATRIX[0].length, sheetContent.getColumnCount());
        }

        @Test
        public void isEmpty() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertFalse(sheetContent.isEmpty());
        }

        @Test
        public void getMatrix() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertArrayEquals(INPUT_MATRIX, sheetContent.getMatrix());
        }

        @Test
        public void getRows() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_LIST, sheetContent.getRows());
        }

        @Test
        public void getRow() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_LIST.get(1), sheetContent.getRow(1));
        }

        @Test
        public void getCell() {
            SheetContent sheetContent = createDefaultSheetContent();
            assertEquals(INPUT_MATRIX[1][1], sheetContent.getCellValue(1, 1));
        }
    }

    @Nested
    @DisplayName("Default Empty Get Behavior")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class EmptyGetTests {

        @Test
        public void getEmptySheetName() {
            SheetContent sheetContent = createEmptySheetContent();
            assertEquals("", sheetContent.getSheetName());
        }

        @Test
        public void getEmptyRowCount() {
            SheetContent sheetContent = createEmptySheetContent();
            assertEquals(0, sheetContent.getRowCount());
        }

        @Test
        public void getEmptyColumnCount() {
            SheetContent sheetContent = createEmptySheetContent();
            assertEquals(0, sheetContent.getColumnCount());
        }

        @Test
        public void isEmptyTrue() {
            SheetContent sheetContent = createEmptySheetContent();
            assertTrue(sheetContent.isEmpty());
        }

        @Test
        public void emptyGetRows() {
            SheetContent sheetContent = createEmptySheetContent();
            List<List<String>> rows = sheetContent.getRows();
            assertNotNull(rows);
            assertEquals(0, rows.size());
        }

        @Test
        public void emptyGetMatrix() {
            SheetContent sheetContent = createEmptySheetContent();
            String[][] matrix = sheetContent.getMatrix();
            assertNotNull(matrix);
            assertArrayEquals(new String[0][0], matrix);
        }
    }

    @Nested
    @DisplayName("Default GET Out Of Range Exception Behavior")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class DefaultGetOutOfRangeTests {
        @Test
        public void getRowIndexNegative() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getRow(-1);
            });
            assertEquals("Row index out of range: -1, size: 2", exception.getMessage());
        }

        @Test
        public void getRowIndexTooLarge() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getRow(200);
            });
            assertEquals("Row index out of range: 200, size: 2", exception.getMessage());
        }

        @Test
        public void getCellWithRowIndexNegative() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getCellValue(-1, 1);
            });
            assertEquals("Row index out of range: -1, size: 2", exception.getMessage());
        }

        @Test
        public void getCellWithRowIndexTooLarge() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getCellValue(9999, 1);
            });
            assertEquals("Row index out of range: 9999, size: 2", exception.getMessage());
        }

        @Test
        public void getCellWithColumnIndexNegative() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getCellValue(1, -1);
            });
            assertEquals("Column index out of range: -1, size: 3", exception.getMessage());
        }

        @Test
        public void getCellWithColumnIndexTooLarge() {
            SheetContent sheetContent = createDefaultSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                sheetContent.getCellValue(1, 9999);
            });
            assertEquals("Column index out of range: 9999, size: 3", exception.getMessage());
        }
    }

    @Test
    public void attemptGetRowAddition() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<String> row = sheetContent.getRow(0);
            //row.set(0, "new value");
            row.add("extra value");
        });
    }

    @Test
    public void attemptGetRowsAddition() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<List<String>> rows = sheetContent.getRows();
            rows.add(List.of("extra value"));
        });
    }
}