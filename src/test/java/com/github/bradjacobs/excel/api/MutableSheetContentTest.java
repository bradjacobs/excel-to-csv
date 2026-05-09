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
import org.junit.jupiter.params.provider.EnumSource;
import org.junit.jupiter.params.provider.MethodSource;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
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

    @ParameterizedTest(name = "Add Row: {0}")
    @MethodSource("listProvider")
    void addRowOperations(String name, List<String> inputList, List<String> expectedList) {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.addRow(inputList);

        List<List<String>> expected = copyOfInput();
        expected.add(expectedList);
        assertEquals(expected, mutableSheetContent.getRows());
    }

    @ParameterizedTest(name = "Insert Row: {0}")
    @MethodSource("listProvider")
    void insertRowOperations(String name, List<String> inputList, List<String> expectedList) {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.insertRow(1, inputList);

        List<List<String>> expected = copyOfInput();
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

    @ParameterizedTest(name = "Replace Row: {0}")
    @MethodSource("listProvider")
    void replaceRowOperations(String name, List<String> inputList, List<String> expectedList) {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.replaceRow(1, inputList);

        List<List<String>> expected = copyOfInput();
        expected.set(1, expectedList);
        assertEquals(expected, mutableSheetContent.getRows());
    }

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
    public void setCellValue() {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        String cellValue = mutableSheetContent.getCellValue(1, 1);
        assertEquals(INPUT_MATRIX[1][1], cellValue);

        String newCellValue = "UPDATED";
        mutableSheetContent.setCellValue(1, 1, newCellValue);
        cellValue = mutableSheetContent.getCellValue(1, 1);
        assertEquals(newCellValue, cellValue);
    }

    @Test
    public void setCellValueWithNull() {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.setCellValue(1, 1, null);
        assertEquals("", mutableSheetContent.getCellValue(1, 1));
    }

    @Test
    public void transformCellValues() {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.transformCells(String::toUpperCase);

        List<List<String>> expected = copyOfInput();
        for (List<String> row : expected) {
            row.replaceAll(String::toUpperCase);
        }
        assertEquals(expected, mutableSheetContent.getRows());
    }

    // Will allow to update row values without having to call 'setRow'
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

    // Will allow to update row values without having to call 'setRow'
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

    enum RowOperation {
        ADD, INSERT, REPLACE, REMOVE, TRANSFORM
    }

    @Nested
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class ExceptionRowOperationsTests {

        @Test
        public void transformRowMissingTransformer() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                mutableSheetContent.transformRow(0, null);
            });
            assertEquals("transformer must not be null", exception.getMessage());
        }

        @Test
        public void transformCellsMissingTransformer() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IllegalArgumentException.class, () -> {
                mutableSheetContent.transformCells(null);
            });
            assertEquals("transformer must not be null", exception.getMessage());
        }

        @Test
        public void getRowsSetRow() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            List<List<String>> rows = mutableSheetContent.getRows();
            assertThrows(UnsupportedOperationException.class, () -> {
                rows.set(0, List.of(""));
            });
        }

        @ParameterizedTest
        @EnumSource(value = RowOperation.class, names = {"INSERT", "REPLACE", "REMOVE", "TRANSFORM"})
        public void testNegativeRowIndexException(RowOperation rowOperation) {
            Executable executable = createExecutable(rowOperation, -1, List.of("foo", "bar", "baz"));
            Exception exception = assertThrows(IndexOutOfBoundsException.class, executable);
            assertEquals("Row index out of range: -1, size: 2", exception.getMessage());
        }

        @ParameterizedTest
        @EnumSource(value = RowOperation.class, names = {"INSERT", "REPLACE", "REMOVE", "TRANSFORM"})
        public void testTooBigRowIndexException(RowOperation rowOperation) {
            Executable executable = createExecutable(rowOperation, 200, List.of("foo", "bar", "baz"));
            Exception exception = assertThrows(IndexOutOfBoundsException.class, executable);
            assertEquals("Row index out of range: 200, size: 2", exception.getMessage());
        }

        // TODO - need new replacement tests - scenario no longer an error.
//        @ParameterizedTest
//        @EnumSource(value = RowOperation.class, names = {"ADD", "INSERT", "REPLACE"})
//        public void testRowTooLargeException(RowOperation rowOperation) {
//            Executable executable = createExecutable(rowOperation, 1, List.of("a", "b", "c", "d", "e", "f", "g"));
//            Exception exception = assertThrows(IllegalArgumentException.class, executable);
//            assertEquals("Row contains too many columns: 7 > 3", exception.getMessage());
//        }

        @Test
        public void setCellWithRowIndexNegative() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.setCellValue(-1, 1, "abc");
            });
            assertEquals("Row index out of range: -1, size: 2", exception.getMessage());
        }

        @Test
        public void setCellWithRowIndexTooLarge() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.setCellValue(9999, 1, "abc");
            });
            assertEquals("Row index out of range: 9999, size: 2", exception.getMessage());
        }

        @Test
        public void setCellWithColumnIndexNegative() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.setCellValue(1, -1, "abc");
            });
            assertEquals("Column index out of range: -1, size: 3", exception.getMessage());
        }

        @Test
        public void setCellWithColumnIndexTooLarge() {
            MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
            Exception exception = assertThrows(IndexOutOfBoundsException.class, () -> {
                mutableSheetContent.setCellValue(1, 9999, "abc");
            });
            assertEquals("Column index out of range: 9999, size: 3", exception.getMessage());
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

    @Test
    public void addColumn() {
        MutableSheetContent mutableSheetContent = createDefaultMutableSheetContent();
        mutableSheetContent.addColumn();
        assertEquals(INPUT_LIST.size(), mutableSheetContent.getRowCount());
        assertEquals(INPUT_LIST.get(0).size() + 1, mutableSheetContent.getColumnCount());

        List<String> newRow = List.of("foo", "bar", "baz", "zed");
        mutableSheetContent.addRow(newRow);

        List<List<String>> expected = copyOfInput();
        for (List<String> exRow : expected) {
            exRow.add("");
        }
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
        for (List<String> exRow : expected) {
            exRow.add(fillerValue);
        }
        expected.add(newRow);
        assertEquals(expected, mutableSheetContent.getRows());
    }

    // TODO -
    //   add test cases for remove columns
}
