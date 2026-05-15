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

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class ByPrefixSheetSelectorTest {

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
            ByPrefixSheetSelector selector = new ByPrefixSheetSelector(List.of("v1"));

            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("v1_Data", 0),
                    sheet("V1_Notes", 1),
                    sheet("v2_Data", 2),
                    sheet("V2_Notes", 3)
            ));

            assertEquals(2, result.size());
            assertEquals("v1_Data", result.get(0).getName());
            assertEquals("V1_Notes", result.get(1).getName());
        }

        @Test
        void normalizesVarargsRequestedNamesCaseInsensitive() {
            ByPrefixSheetSelector selector = new ByPrefixSheetSelector("c", "D");

            List<SheetInfo> result = selector.filterSheets(sheets(
                    sheet("cat", 0),
                    sheet("cow", 1),
                    sheet("dog", 2),
                    sheet("rat", 3)
            ));

            assertEquals(3, result.size());
            assertEquals("cat", result.get(0).getName());
            assertEquals("cow", result.get(1).getName());
            assertEquals("dog", result.get(2).getName());
        }

        @Nested
        @DisplayName("exception behavior")
        class ExceptionTests {
            @Test
            void throwsWithNullValues() {
                Collection<String> values = null;
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector(values));
                assertEquals("Prefixes cannot be null", e.getMessage());
            }

            @Test
            void throwsWithNullVarArgs() {
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector((String[]) null));
                assertEquals("Prefixes cannot be null", e.getMessage());
            }

            @Test
            void throwsWithEmptyValues() {
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector(List.of()));
                assertEquals("Prefixes cannot be empty", e.getMessage());
            }

            @Test
            void throwsWhenValuesContainNull() {
                List<String> values = new ArrayList<>();
                values.add(null);
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector(values));
                assertEquals("Prefixes cannot contain null values", e.getMessage());
            }

            @Test
            void throwsWhenVarArgsContainNull() {
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector("foo", null, "bar"));
                assertEquals("Prefixes cannot contain null values", e.getMessage());
            }

            @Test
            void throwsWithValuesContainEmptyString() {
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector(List.of("foo", "")));
                assertEquals("Prefixes cannot contain empty values", e.getMessage());
            }

            @Test
            void throwsWithDuplicateValues() {
                Exception e = assertThrows(IllegalArgumentException.class,
                        () -> new ByPrefixSheetSelector(List.of("one", "two", "ONE")));
                assertEquals("Prefixes cannot contain duplicate values", e.getMessage());
            }
        }
    }
}
