/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

public class BasicSheetContentTest extends SheetContentTest {

    private static final SheetContent DEFAULT_INPUT_SHEET = new BasicSheetContent(INPUT_SHEET_NAME, INPUT_MATRIX);
    private static final SheetContent EMPTY_INPUT_SHEET = new BasicSheetContent("", new String[][]{});

    @Override
    protected SheetContent createDefaultSheetContent() {
        return DEFAULT_INPUT_SHEET;
    }

    @Override
    protected SheetContent createEmptySheetContent() {
        return EMPTY_INPUT_SHEET;
    }

    @Test
    public void testWithNullDataInput() {
        SheetContent sheetContent = new BasicSheetContent("mySheet", null);
        assertEquals(0, sheetContent.getRowCount());
        assertEquals(0, sheetContent.getColumnCount());
        assertNotNull(sheetContent.getRows());
        assertEquals(0, sheetContent.getRows().size());
    }

    @Test
    public void testWithNullSheetNameInput() {
        SheetContent sheetContent = new BasicSheetContent(null, INPUT_MATRIX);
        assertEquals("", sheetContent.getSheetName());
    }

    @Test
    public void attemptUpdateGetRow() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<String> row = sheetContent.getRow(0);
            row.set(0, "new value");
        });
    }

    @Test
    public void attemptUpdateGetRows() {
        SheetContent sheetContent = createDefaultSheetContent();
        assertThrows(UnsupportedOperationException.class, () -> {
            List<List<String>> rows = sheetContent.getRows();
            rows.set(0, List.of("new value"));
        });
    }
}
