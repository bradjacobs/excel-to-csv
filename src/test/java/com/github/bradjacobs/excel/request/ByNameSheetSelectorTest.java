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
class ByNameSheetSelectorTest {

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
        void normalizesRequestedNamesCaseInsensitive() {
            ByNameSheetSelector selector = new ByNameSheetSelector(Arrays.asList("ALPHA", "Beta"));

            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("alpha", 0),
                    sheet("beta", 1),
                    sheet("gamma", 2)
            ));

            assertEquals(2, result.size());
            assertEquals("alpha", result.get(0).getName());
            assertEquals("beta", result.get(1).getName());
        }

        @Test
        void normalizesVarargsRequestedNamesCaseInsensitive() {
            ByNameSheetSelector selector = new ByNameSheetSelector("ALPHA", "Beta");

            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("alpha", 0),
                    sheet("beta", 1),
                    sheet("gamma", 2)
            ));

            assertEquals(2, result.size());
            assertEquals("alpha", result.get(0).getName());
            assertEquals("beta", result.get(1).getName());
        }

        @Test
        void returnsRequestedSheetsInRequestedOrder() {
            ByNameSheetSelector selector = new ByNameSheetSelector(Arrays.asList("gamma", "alpha"));

            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("Alpha", 0),
                    sheet("Beta", 1),
                    sheet("Gamma", 2)
            ));

            assertEquals(2, result.size());
            assertEquals("Gamma", result.get(0).getName());
            assertEquals("Alpha", result.get(1).getName());
        }
    }

    @Nested
    @DisplayName("exception behavior")
    class ExceptionTests {
        @Test
        void throwsWithNullValues() {
            Collection<String> values = null;
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector(values));
            assertEquals("Names cannot be null", e.getMessage());
        }

        @Test
        void throwsWithNullVarArgs() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector((String[])null));
            assertEquals("Names cannot be null", e.getMessage());
        }

        @Test
        void throwsWithEmptyValues() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector(List.of()));
            assertEquals("Names cannot be empty", e.getMessage());
        }

        @Test
        void throwsWhenValuesContainNull() {
            List<String> values = new ArrayList<>();
            values.add(null);
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector(values));
            assertEquals("Names cannot have null values", e.getMessage());
        }

        @Test
        void throwsWhenVarArgsContainNull() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector("foo", null, "bar"));
            assertEquals("Names cannot have null values", e.getMessage());
        }

        @Test
        void throwsWithValuesContainEmptyString() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector(List.of("foo", "")));
            assertEquals("Names cannot contain empty values", e.getMessage());
        }

        @Test
        void throwsWhenRequestedNameMissing() {
            // NOTE:  want the throw message to have sheetName case that was passed in.
            ByNameSheetSelector selector = new ByNameSheetSelector(List.of("one", "FAKE"));
            IllegalArgumentException e = assertThrows(IllegalArgumentException.class,
                    () -> selector.filterSheets(sheets(
                            sheet("zero", 0),
                            sheet("one", 1)
                    )));
            assertEquals("Requested Excel sheet not found: 'FAKE'", e.getMessage());
        }

        @Test
        void throwsWithDuplicateValues() {
            Exception e = assertThrows(IllegalArgumentException.class,
                    () -> new ByNameSheetSelector(List.of("one", "two", "ONE")));
            assertEquals("Names cannot contain duplicate values", e.getMessage());
        }

        @Test
        void throwsWhenSheetsContainsDuplicateIndexes() {
            ByNameSheetSelector selector = new ByNameSheetSelector(List.of("foo"));
            IllegalStateException e = assertThrows(IllegalStateException.class,
                    // _NOTE_: this won't happen with a real excel file.
                    () -> selector.filterSheets(sheets(
                            sheet("foo", 0),
                            sheet("foo", 0)
                    )));
            assertTrue(e.getMessage().contains("Duplicate key"));
        }
    }
}
