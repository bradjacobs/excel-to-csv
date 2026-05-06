/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import com.github.bradjacobs.excel.util.TestSheetInfoUtil;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class ByIndexSheetSelectorTest {

    private static SheetInfo sheet(String name, int index) {
       return TestSheetInfoUtil.sheet(name, index);
    }

    private static List<SheetInfo> sheets(SheetInfo... sheetInfos) {
        return Arrays.asList(sheetInfos);
    }

    @Nested
    @DisplayName("selection behavior")
    class SelectionTests {
        @Test
        void returnsRequestedSheetsInRequestedOrder() {
            ByIndexSheetSelector selector = new ByIndexSheetSelector(2, 0);
            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("zero", 0),
                    sheet("one", 1),
                    sheet("two", 2)
            ));
            assertEquals(2, result.size(), "Mismatch expected sheet count");
            assertEquals(2, result.get(0).getIndex());
            assertEquals(0, result.get(1).getIndex());
        }
    }

    @Nested
    @DisplayName("exception behavior")
    class ExceptionTests {
        @Test
        void throwsWithNullValues() {
            Collection<Integer> values = null;
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector(values));
            assertEquals("Indexes cannot be null", e.getMessage());
        }

        @Test
        void throwsWithNullVarArgs() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector((int[])null));
            assertEquals("Indexes cannot be null", e.getMessage());
        }

        @Test
        void throwsWithEmptyValues() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector(List.of()));
            assertEquals("Indexes cannot be empty", e.getMessage());
        }

        @Test
        void throwsWhenValuesContainNull() {
            List<Integer> values = new ArrayList<>();
            values.add(null);
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector(values));
            assertEquals("Indexes cannot contain null values", e.getMessage());
        }

        @Test
        void throwsWithDuplicateValues() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector(List.of(3, 1, 3)));
            assertEquals("Indexes cannot contain duplicate values", e.getMessage());
        }

        @Test
        void throwsWithNegativeValues() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByIndexSheetSelector(List.of(-2)));
            assertEquals("Indexes cannot contain negative values", e.getMessage());
        }

        @Test
        void throwsWhenRequestedIndexMissing() {
            ByIndexSheetSelector selector = new ByIndexSheetSelector(List.of(0, 99));
            IllegalArgumentException e = assertThrows(IllegalArgumentException.class,
                    () -> selector.filterSheets(sheets(
                            sheet("zero", 0),
                            sheet("one", 1)
                    )));
            assertEquals("Requested Excel sheet not found: '99'", e.getMessage());
        }

        @Test
        void throwsWhenSheetsContainsDuplicateIndexes() {
            ByIndexSheetSelector selector = new ByIndexSheetSelector(List.of(0));
            IllegalStateException e = assertThrows(IllegalStateException.class,
                    // _NOTE_: this won't happen with a real excel file.
                    () -> selector.filterSheets(sheets(
                            sheet("a", 0),
                            sheet("b", 0)
                    )));
            assertTrue(e.getMessage().contains("Duplicate key"));
        }
    }
}
