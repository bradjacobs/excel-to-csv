/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.api.function.Executable;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.CsvSource;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.params.provider.Arguments.arguments;

class MutableSheetContentTest extends SheetContentTest {

    private static final SheetContent INPUT_SHEET = new BasicSheetContent(INPUT_SHEET_NAME, INPUT_MATRIX);

    @Override
    protected SheetContent createDefaultSheetContent() {
        return MutableSheetContent.copyOf(INPUT_SHEET);
    }

    private MutableSheetContent createDefaultMutableSheetContent() {
        return MutableSheetContent.copyOf(INPUT_SHEET);
    }

    @Override
    protected SheetContent createEmptySheetContent() {
        SheetContent emptySheet = new BasicSheetContent("", new String[][]{});
        return MutableSheetContent.copyOf(emptySheet);
    }

    static Stream<Arguments> listProvider() {
        return Stream.of(
                row("simple row", List.of("foo","bar","baz"), List.of("foo","bar","baz")),
                row("row contains a null", list("foo", null, "baz"), List.of("foo","", "baz")),
                row("row too short", List.of("foo"), List.of("foo","","")),
                row("row long with blanks", List.of("cat","dog","rat","",""), List.of("cat","dog","rat")),
                row("row long expands columns", List.of("cat","dog","rat","cow","hen"), List.of("cat","dog","rat","cow","hen")),
                row("row with leading blank", List.of("", "cow", "cat"), List.of("","cow","cat")),
                row("row with special characters", List.of("Façade", "\"cat's\"", "[dogs]{}"), List.of("Façade", "\"cat's\"", "[dogs]{}")),
                row("null input row list", null, List.of("","","")),
                row("empty input row list", List.of(), List.of("","",""))
        );
    }

    static Arguments row(String name, List<String> in, List<String> out) {
        return arguments(name, in, out);
    }

    static List<String> list(String... values) {
        return new ArrayList<>(Arrays.asList(values)); // allows nulls
    }

    enum RowOperation {
        ADD, INSERT, REPLACE, REMOVE, TRANSFORM
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class SetSheetNameTests {
        @Test
        public void setSheetName() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            String newSheetName = "newSheetName";
            mutableSheetContent.setSheetName(newSheetName);
            assertEquals(newSheetName, mutableSheetContent.getSheetName());
        }

        @Test
        public void setSheetNameWithNull() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.setSheetName(null);
            assertEquals("", mutableSheetContent.getSheetName());
        }

        @Test
        public void createWithNullSheetName() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent(null, INPUT_LIST);
            assertEquals("", mutableSheetContent.getSheetName());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class CreateWithEmptyDataTests {

        @Test
        public void createMutableSheetContentWithNullRows() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("sheet", null);
            assertEquals(0, mutableSheetContent.getRowCount());
            assertEquals(0, mutableSheetContent.getColumnCount());
            assertTrue(mutableSheetContent.isEmpty());

            List<List<String>> rows = mutableSheetContent.getRows();
            assertNotNull(rows);
            assertEquals(0, rows.size());
        }

        @Test
        public void copyOfNullSheetContent() {
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                MutableSheetContent.copyOf(null);
            });
            assertEquals("sheetContent cannot be null", exception.getMessage());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AddRowTests {

        @ParameterizedTest(name = "Add Row: {0}")
        @MethodSource("com.github.bradjacobs.excel.api.MutableSheetContentTest#listProvider")
        void addRowOperations(String name, List<String> inputList, List<String> expectedList) {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.addRow(inputList);

            List<String> fetchedRow = mutableSheetContent.getRow(mutableSheetContent.getRowCount() - 1);
            assertEquals(expectedList, fetchedRow);

            List<List<String>> expected = copyOfInput(expectedList.size());
            expected.add(expectedList);
            assertEquals(expected, mutableSheetContent.getRows());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class InsertRowTests {

        @ParameterizedTest(name = "Insert Row: {0}")
        @MethodSource("com.github.bradjacobs.excel.api.MutableSheetContentTest#listProvider")
        void insertRowOperations(String name, List<String> inputList, List<String> expectedList) {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.insertRow(1, inputList);

            List<String> fetchedRow = mutableSheetContent.getRow(1);
            assertEquals(expectedList, fetchedRow);

            List<List<String>> expected = copyOfInput(expectedList.size());
            expected.add(1, expectedList);
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void insertRowInFront() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<String> newRow = List.of("foo", "bar", "baz");
            mutableSheetContent.insertRow(0, newRow);

            List<List<String>> expected = copyOfInput();
            expected.add(0, newRow);
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void insertRowAtEnd() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<String> newRow = List.of("foo", "bar", "baz");
            mutableSheetContent.insertRow(INPUT_LIST.size(), newRow);

            List<List<String>> expected = copyOfInput();
            expected.add(newRow);
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void insertRowNegativeRowIndexException() {
            runNegativeRowIndexExceptionScenario(RowOperation.INSERT);
        }

        @Test
        public void insertRowTooBigRowIndexException() {
            runTooBigRowIndexExceptionScenario(RowOperation.INSERT);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class ReplaceRowTests {

        @ParameterizedTest(name = "Replace Row: {0}")
        @MethodSource("com.github.bradjacobs.excel.api.MutableSheetContentTest#listProvider")
        void replaceRowOperations(String name, List<String> inputList, List<String> expectedList) {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.replaceRow(1, inputList);

            List<String> fetchedRow = mutableSheetContent.getRow(1);
            assertEquals(expectedList, fetchedRow);

            List<List<String>> expected = copyOfInput(expectedList.size());
            expected.set(1, expectedList);
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void replaceRowNegativeRowIndexException() {
            runNegativeRowIndexExceptionScenario(RowOperation.REPLACE);
        }

        @Test
        public void replaceRowTooBigRowIndexException() {
            runTooBigRowIndexExceptionScenario(RowOperation.REPLACE);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class TransformRowTests {

        @Test
        public void transformRow() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.transformRow(0, String::toUpperCase);

            List<List<String>> expected = copyOfInput();
            List<String> expectedFirstRow = expected.get(0);
            expectedFirstRow.replaceAll(String::toUpperCase);
            assertEquals(expected, mutableSheetContent.getRows());

            assertEquals(expectedFirstRow, mutableSheetContent.getRow(0));
        }

        @Test
        public void transformRowMissingTransformer() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                mutableSheetContent.transformRow(0, null);
            });
            assertEquals("transformer must not be null", exception.getMessage());
        }

        @Test
        public void transformRowNegativeRowIndexException() {
            runNegativeRowIndexExceptionScenario(RowOperation.TRANSFORM);
        }

        @Test
        public void transformRowTooBigRowIndexException() {
            runTooBigRowIndexExceptionScenario(RowOperation.TRANSFORM);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class RemoveRowTests {
        @Test
        public void removeRow() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<String> removedRow = mutableSheetContent.removeRow(0);

            List<List<String>> expected = copyOfInput();
            List<String> expectedRemovedRow = expected.remove(0);
            assertEquals(expectedRemovedRow, removedRow);
            assertEquals(expected, mutableSheetContent.getRows());

            List<String> theNowFirstRow = mutableSheetContent.getRow(0);
            assertEquals(expected.get(0), theNowFirstRow);
        }

        @Test
        public void removeAllRows() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeRow(0);
            mutableSheetContent.removeRow(0);

            assertEquals(0, mutableSheetContent.getRowCount(), "mismatch expected row count");
            assertEquals(0, mutableSheetContent.getColumnCount(), "mismatch expected column count");
            assertTrue(mutableSheetContent.isEmpty());
        }

        @Test
        public void removeRowNegativeRowIndexException() {
            runNegativeRowIndexExceptionScenario(RowOperation.REMOVE);
        }

        @Test
        public void removeRowTooBigRowIndexException() {
            runTooBigRowIndexExceptionScenario(RowOperation.REMOVE);
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class SetCellTests {
        @Test
        public void setCellValue() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            String newCellValue = "UPDATED";
            mutableSheetContent.setCellValue(1, 1, newCellValue);
            String cellValue = mutableSheetContent.getCellValue(1, 1);
            assertEquals(newCellValue, cellValue);
        }

        @Test
        public void setCellValueWithNull() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.setCellValue(1, 1, null);
            assertEquals("", mutableSheetContent.getCellValue(1, 1));
        }

        @ParameterizedTest
        @CsvSource({
                "-1,1,'Row index out of range: -1, size: 2'",
                "9,1,'Row index out of range: 9, size: 2'",
                "1,-1,'Column index out of range: -1, size: 3'",
                "1,9,'Column index out of range: 9, size: 3'",
        })
        void setCellValueInvalidIndex(Integer rowIndex, Integer columnIndex, String expectedError) {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.setCellValue(rowIndex, columnIndex, "abc");
            });
            assertEquals(expectedError, exception.getMessage());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class TransformCellTests {

        @Test
        public void transformCellValues() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.transformCells(String::toUpperCase);

            List<List<String>> expected = copyOfInput();
            expected.forEach(exRow -> exRow.replaceAll(String::toUpperCase));
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void transformCellsMissingTransformer() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                mutableSheetContent.transformCells(null);
            });
            assertEquals("transformer must not be null", exception.getMessage());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class AddColumnTests {

        @Test
        public void addColumn() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.addColumn();
            assertEquals(INPUT_LIST.size(), mutableSheetContent.getRowCount());
            assertEquals(INPUT_LIST.get(0).size() + 1, mutableSheetContent.getColumnCount());

            // add a new row with the new column count
            List<String> newRow = List.of("foo", "bar", "baz", "zed");
            mutableSheetContent.addRow(newRow);

            List<List<String>> expected = copyOfInput();
            expected.forEach(exRow -> exRow.add(BLANK));
            expected.add(newRow);
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void addColumnWithFillerValue() {
            String fillerValue = "filler";
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.addColumn(fillerValue);
            assertEquals(INPUT_LIST.size(), mutableSheetContent.getRowCount());
            assertEquals(INPUT_LIST.get(0).size() + 1, mutableSheetContent.getColumnCount());

            List<String> newRow = List.of("foo", "bar", "baz", "zed");
            mutableSheetContent.addRow(newRow);

            List<List<String>> expected = copyOfInput();
            expected.forEach(exRow -> exRow.add(fillerValue));
            expected.add(newRow);
            assertEquals(expected, mutableSheetContent.getRows());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class RemoveColumnTests {
        private final String[][] INPUT_MATRIX_FOR_REMOVE_COLUMN = {
                {"Id", "Name", "Age", "Notes"},
                {"22", "Bo", "22", "foo"},
                {"33", "Flo", "29", "bar"}
        };
        private final List<List<String>> REMOVE_COLUMNS_INPUT_ROWS = matrixToList(INPUT_MATRIX_FOR_REMOVE_COLUMN);

        // tests will remove 2 columns, and below
        // is the expected after 2 columns removed.
        private final String[][] REMOVE_COLUMNS_EXPECTED_MATRIX = {
                {"Id", "Age"},
                {"22", "22"},
                {"33", "29"}
        };
        private final List<List<String>> REMOVE_COLUMNS_EXPECTED_ROWS = matrixToList(REMOVE_COLUMNS_EXPECTED_MATRIX);


        @Test
        public void removeColumnsByIndex() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", REMOVE_COLUMNS_INPUT_ROWS);
            mutableSheetContent.removeColumn(1, 3);
            assertEquals(REMOVE_COLUMNS_EXPECTED_ROWS, mutableSheetContent.getRows());
        }

        @Test
        public void removeColumnsWithDuplicatesByIndex() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", REMOVE_COLUMNS_INPUT_ROWS);
            mutableSheetContent.removeColumn(1, 3, 1, 1, 3);
            assertEquals(REMOVE_COLUMNS_EXPECTED_ROWS, mutableSheetContent.getRows());
        }

        @Test
        public void removeAllColumnsByIndex() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", REMOVE_COLUMNS_INPUT_ROWS);
            mutableSheetContent.removeColumn(0,1,2,3);
            assertEquals(List.of(), mutableSheetContent.getRows());
        }

        @Test
        public void removeColumnsByHeaderName() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", REMOVE_COLUMNS_INPUT_ROWS);
            mutableSheetContent.removeColumn("Name", "Notes");
            assertEquals(REMOVE_COLUMNS_EXPECTED_ROWS, mutableSheetContent.getRows());
        }

        @Test
        public void removeColumnsByHeaderNameCaseInsensitive() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", REMOVE_COLUMNS_INPUT_ROWS);
            mutableSheetContent.removeColumn("name", "NOTES");
            assertEquals(REMOVE_COLUMNS_EXPECTED_ROWS, mutableSheetContent.getRows());
        }

        @Test
        public void removeColumnsUnknownHeaderNames() {
            // trying to delete a header that doesn't exist will have no effect
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeColumn("FOOBAR");
            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void testRemoveNegativeColumnIndexException() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.removeColumn(-1);
            });
            assertEquals("Column index out of range: -1, size: 3", exception.getMessage());
        }

        @Test
        public void testRemoveTooBigColumnIndexException() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.removeColumn(99);
            });
            assertEquals("Column index out of range: 99, size: 3", exception.getMessage());
        }

        @Test
        public void testRemoveColumnWithNameNoRows() {
            MutableSheetContent mutableSheetContent = new MutableSheetContent("", List.of());
            Exception exception = assertThrows(IllegalStateException.class, () -> {
                mutableSheetContent.removeColumn("Name");
            });
            assertEquals("Cannot find column indexes when sheet is empty", exception.getMessage());
        }

        @Test
        public void removeWithInvalidColumnNoSideEffect() {
            // if an invalid index was passed in, ensure that NO columns were removed.
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.removeColumn(0, -8, 1);
            });
            assertEquals("Column index out of range: -8, size: 3", exception.getMessage());

            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void removeNullColumnIndex() {
            // if an invalid index was passed in, ensure that NO columns were removed.
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeColumn((int[])null);
            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void removeEmptyColumnIndex() {
            // if an invalid index was passed in, ensure that NO columns were removed.
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeColumn(new int[0]);
            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void removeNullColumnHeaderNames() {
            // if an invalid index was passed in, ensure that NO columns were removed.
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeColumn((String[])null);
            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }

        @Test
        public void removeEmptyColumnHeaderNames() {
            // if an invalid index was passed in, ensure that NO columns were removed.
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            mutableSheetContent.removeColumn(new String[0]);
            List<List<String>> expected = copyOfInput();
            assertEquals(expected, mutableSheetContent.getRows());
        }
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class GetRowsAndUpdateTests {

        // Will allow updating row values without having to call 'setRow'
        //  (may change mind and not allow later, tbd)
        @Test
        public void getRowUpdateValue() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<String> firstRow = mutableSheetContent.getRow(0);

            String newValue = "UPDATED_VALUE";
            firstRow.set(1, newValue);

            List<String> refetchedFirstRow = mutableSheetContent.getRow(0);
            assertEquals(newValue, refetchedFirstRow.get(1));
            assertEquals(newValue, mutableSheetContent.getCellValue(0,1));

            // try setting to null
            firstRow.set(1, null);
            assertEquals("", mutableSheetContent.getCellValue(0,1));
        }

        // Will allow updating row values without having to call 'setRow'
        //  (may change mind and not allow later, tbd)
        @Test
        public void getRowsUpdateValue() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<List<String>> rows = mutableSheetContent.getRows();

            String newValue = "UPDATED_VALUE";
            List<String> firstRow = rows.get(0);
            firstRow.set(1, newValue);

            List<String> refetchedFirstRow = mutableSheetContent.getRow(0);
            assertEquals(newValue, refetchedFirstRow.get(1));
            assertEquals(newValue, mutableSheetContent.getCellValue(0,1));
        }

        @Test
        public void attemptInsertRowOnGetRows() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<List<String>> rows = mutableSheetContent.getRows();
            assertThrows(UnsupportedOperationException.class, () -> {
                rows.add(0, List.of("temp"));
            });
        }

        @Test
        public void attemptReplaceRowOnGetRows() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<List<String>> rows = mutableSheetContent.getRows();
            assertThrows(UnsupportedOperationException.class, () -> {
                rows.set(0, List.of("temp"));
            });
        }
    }


    private void runNegativeRowIndexExceptionScenario(RowOperation rowOperation) {
        Executable executable = createExecutable(rowOperation, -1, List.of("foo", "bar", "baz"));
        Exception exception = assertThrows(IndexOutOfBoundsException.class, executable);
        assertEquals("Row index out of range: -1, size: 2", exception.getMessage());
    }

    public void runTooBigRowIndexExceptionScenario(RowOperation rowOperation) {
        Executable executable = createExecutable(rowOperation, 200, List.of("foo", "bar", "baz"));
        Exception exception = assertThrows(IndexOutOfBoundsException.class, executable);
        assertEquals("Row index out of range: 200, size: 2", exception.getMessage());
    }

    private Executable createExecutable(RowOperation rowOperation, int rowIndex, List<String> row) {
        MutableSheetContent mutableContent = createDefaultMutableSheetContent();
        switch (rowOperation) {
            case ADD:
                return () -> mutableContent.addRow(row);
            case INSERT:
                return () -> mutableContent.insertRow(rowIndex, row);
            case REPLACE:
                return () -> mutableContent.replaceRow(rowIndex, row);
            case TRANSFORM:
                return () -> mutableContent.transformRow(rowIndex, String::toUpperCase);
            case REMOVE:
                return () -> mutableContent.removeRow(rowIndex);
            default:
                throw new IllegalStateException("Unexpected value: " + rowOperation);
        }
    }
}
