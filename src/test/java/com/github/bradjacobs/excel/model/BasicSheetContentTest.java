/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.model;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

class BasicSheetContentTest extends SheetContentTest {

    private static final SheetContent DEFAULT_INPUT_SHEET = new BasicSheetContent(INPUT_SHEET_NAME, INPUT_MATRIX);
    private static final SheetContent EMPTY_INPUT_SHEET = new BasicSheetContent("", new String[][]{});
    private static final SheetContent WITH_DULL_INPUT_SHEET = new BasicSheetContent("", INPUT_LIST_W_NULL);

    @Override
    protected SheetContent createDefaultSheetContent() {
        return DEFAULT_INPUT_SHEET;
    }

    @Override
    protected SheetContent createEmptySheetContent() {
        return EMPTY_INPUT_SHEET;
    }

    @Override
    protected SheetContent createSheetWithNullContent() {
        return WITH_DULL_INPUT_SHEET;
    }

    @Test
    void testWithNullDataMatrixInput() {

        final String[][] JAGGED_ARRAY = {
                {"A", "B", "C"},
                {"D"},
                {"E", "F"},
                {"G", "H", "I", "J"}
        };

        SheetContent sheetContent2 = new BasicSheetContent("mySheet", JAGGED_ARRAY);


        SheetContent sheetContent = new BasicSheetContent("mySheet", (String[][])null);
        assertEquals(0, sheetContent.getRowCount());
        assertEquals(0, sheetContent.getColumnCount());
        assertNotNull(sheetContent.getRows());
        assertEquals(0, sheetContent.getRows().size());
    }

    @Test
    void testWithNullDataRowsInput() {
        SheetContent sheetContent = new BasicSheetContent("mySheet", (List<List<String>>)null);
        assertEquals(0, sheetContent.getRowCount());
        assertEquals(0, sheetContent.getColumnCount());
        assertNotNull(sheetContent.getRows());
        assertEquals(0, sheetContent.getRows().size());
    }

    @Test
    void testWithNullSheetNameInput() {
        SheetContent sheetContent = new BasicSheetContent(null, INPUT_MATRIX);
        assertEquals("", sheetContent.getSheetName());
    }

    @Test
    void attemptUpdateGetRow() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<String> row = sheetContent.getRow(0);
            row.set(0, "new value");
        });
    }

    @Test
    void attemptUpdateGetRows() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<List<String>> rows = sheetContent.getRows();
            rows.set(0, List.of("new value"));
        });
    }

}
