/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.row;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import org.apache.commons.collections4.list.UnmodifiableList;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RowDataUtilTest {

    @Nested
    class ToUnmodifiableRowsTests {

        @Test
        void arrayNullToRows() {
            List<List<String>> result = RowDataUtil.toUnmodifiableRows((String[][]) null);
            assertUnmodifiableRows(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void arrayEmptyToRows() {
            List<List<String>> result = RowDataUtil.toUnmodifiableRows(new String[0][0]);
            assertUnmodifiableRows(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void listNullToRows() {
            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows((List<List<String>>) null);
            assertUnmodifiableRows(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void normalArrayDataToRows() {
            String[][] input = {
                    {"A", "B"},
                    {"C", "D"}
            };

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(input);

            assertUnmodifiableRows(result);
            assertEquals(2, result.size());
            assertEquals(List.of("A", "B"), result.get(0));
            assertEquals(List.of("C", "D"), result.get(1));
        }

        @Test
        void arrayContainsRowWithNullElementsToRows() {
            // NOTE: this specific util is not responsible
            //   for converting null to empty string for
            //   individual values in the list.
            String[][] input = {
                    {"A", null, "C"}
            };

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(input);

            assertUnmodifiableRows(result);
            assertEquals(1, result.size());
            assertEquals(3, result.get(0).size());
            assertEquals("A", result.get(0).get(0));
            assertNull(result.get(0).get(1));
            assertEquals("C", result.get(0).get(2));
        }

        @Test
        void arrayWithNullRowToRows() {
            // a "null row" will become an "empty list row"
            String[][] input = { null };

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(input);

            assertUnmodifiableRows(result);
            assertEquals(1, result.size());
            assertTrue(result.get(0).isEmpty());
        }

        @Test
        void rowsReturnsSameInstanceWhenAlreadyUnmodifiable() {
            List<List<String>> original =
                    new UnmodifiableList<>(List.of(List.of("A")));

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(original);
            assertSame(original, result);
        }

        @Test
        void normalListDataToRows() {
            List<List<String>> input = new ArrayList<>();
            input.add(Arrays.asList("A", "B"));
            input.add(List.of("C"));

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(input);

            assertUnmodifiableRows(result);
            assertEquals(2, result.size());
            assertEquals(List.of("A", "B"), result.get(0));
            assertEquals(List.of("C"), result.get(1));
        }

        @Test
        void inputListWithEmptyNestedRowToRows() {
            List<List<String>> input = List.of(List.of());

            List<List<String>> result =
                    RowDataUtil.toUnmodifiableRows(input);

            assertUnmodifiableRows(result);
            assertEquals(1, result.size());
            assertTrue(result.get(0).isEmpty());
        }
    }

    @Nested
    class ToUnmodifiableRowTests {

        @Test
        void arrayNullToRow() {
            List<String> result =
                    RowDataUtil.toUnmodifiableRow((String[]) null);
            assertUnmodifiableRow(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void arrayEmptyToRow() {
            List<String> result =
                    RowDataUtil.toUnmodifiableRow(new String[0]);
            assertUnmodifiableRow(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void listNullToRow() {
            List<String> result =
                    RowDataUtil.toUnmodifiableRow((List<String>) null);
            assertUnmodifiableRow(result);
            assertTrue(result.isEmpty());
        }

        @Test
        void normalArrayDataToRow() {
            List<String> result =
                    RowDataUtil.toUnmodifiableRow(
                            new String[]{"A", "B", "C"});

            assertUnmodifiableRow(result);
            assertEquals(List.of("A", "B", "C"), result);
        }

        @Test
        void arrayWithNullValueToRow() {
            // NOTE: this specific util is not responsible
            //   for converting null to empty string for
            //   individual values in the list.
            String[] input = {"A", null, "C"};

            List<String> result =
                    RowDataUtil.toUnmodifiableRow(input);

            assertUnmodifiableRow(result);
            assertEquals(3, result.size());
            assertEquals("A", result.get(0));
            assertNull(result.get(1));
            assertEquals("C", result.get(2));
        }

        @Test
        void rowReturnsSameInstanceWhenAlreadyUnmodifiable() {
            List<String> original =
                    new UnmodifiableList<>(List.of("A"));
            List<String> result =
                    RowDataUtil.toUnmodifiableRow(original);
            assertSame(original, result);
        }

        @Test
        void normalListDataToRow() {
            List<String> input =
                    new ArrayList<>(List.of("A", "B"));

            List<String> result =
                    RowDataUtil.toUnmodifiableRow(input);

            assertUnmodifiableRow(result);
            assertEquals(List.of("A", "B"), result);
        }

        @Test
        void listWithNullValueToRow() {
            // NOTE: this specific util is not responsible
            //   for converting null to empty string for
            //   individual values in the list.
            List<String> input =
                    Arrays.asList("A", null, "C");

            List<String> result =
                    RowDataUtil.toUnmodifiableRow(input);

            assertUnmodifiableRow(result);
            assertEquals(3, result.size());
            assertEquals("A", result.get(0));
            assertNull(result.get(1));
            assertEquals("C", result.get(2));
        }
    }

    @Nested
    class To2DArrayTests {

        @Test
        void nullListInputToEmpty2DArray() {
            String[][] result =
                    RowDataUtil.toArray(null);
            assertNotNull(result);
            assertEquals(0, result.length);
        }

        @Test
        void emptyListInputToEmpty2DArray() {
            String[][] result =
                    RowDataUtil.toArray(List.of());
            assertNotNull(result);
            assertEquals(0, result.length);
        }

        @Test
        void normalRowsTo2DArray() {
            List<List<String>> rows = List.of(
                    List.of("A", "B"),
                    List.of("C", "D")
            );

            String[][] result =
                    RowDataUtil.toArray(rows);

            assertNotNull(result);
            assertEquals(2, result.length);
            assertArrayEquals(
                    new String[]{"A", "B"},
                    result[0]);

            assertArrayEquals(
                    new String[]{"C", "D"},
                    result[1]);
        }

        @Test
        void listWithEmptyRowTo2DArray() {
            List<List<String>> rows = List.of(
                    List.of()
            );

            String[][] result =
                    RowDataUtil.toArray(rows);

            assertNotNull(result);
            assertEquals(1, result.length);
            assertEquals(0, result[0].length);
        }

        @Test
        void listWithRowContainsNullValueTo2DArray() {
            List<List<String>> rows = List.of(
                    Arrays.asList("A", null, "C")
            );

            String[][] result =
                    RowDataUtil.toArray(rows);

            assertNotNull(result);
            assertArrayEquals(
                    new String[]{"A", null, "C"},
                    result[0]);
        }
    }

    private void assertUnmodifiableRows(List<List<String>> result) {
        assertNotNull(result, "Expected non-null result");
        assertThrows(UnsupportedOperationException.class,
                () -> result.add(List.of()));

        if (!result.isEmpty()) {
            assertThrows(UnsupportedOperationException.class,
                    () -> result.set(0, List.of("Q")));
            assertUnmodifiableRow(result.get(0));
        }
    }

    void assertUnmodifiableRow(List<String> result) {
        assertNotNull(result, "Expected non-null result");
        assertThrows(UnsupportedOperationException.class,
                () -> result.add("A"));

        if (!result.isEmpty()) {
            assertThrows(UnsupportedOperationException.class,
                    () -> result.set(0, "X"));
        }
    }
}
