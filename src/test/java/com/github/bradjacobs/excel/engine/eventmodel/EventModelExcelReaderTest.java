/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel;

import com.github.bradjacobs.excel.engine.AbstractExcelReaderTest;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class EventModelExcelReaderTest extends AbstractExcelReaderTest<EventModelExcelReader, EventModelExcelReader.Builder> {

    @Override
    protected EventModelExcelReader.Builder createBuilder() {
        return EventModelExcelReader.builder();
    }

    @Test
    void oldFileFormatThrowsException() {
        // event reader doesn't support older formats,
        // thus expect exception thrown if given old file format.
        Path oldFilePath = TestResourceUtil.getResourceFilePath("test_data.xls");

        Exception thrown = assertThrows(IOException.class, () -> {
            EventModelExcelReader reader = createBuilder().build();
            ExcelReadRequest request = ExcelReadRequest.from(oldFilePath).build();
            reader.readSheet(request);
        });
        assertEquals("Failed to read Excel sheet data: " +
                        "Cannot open workbook - unsupported file type: OLE2",
                thrown.getMessage());
    }

    // TODO
    //   1. add test when manually create new SharedStringsTable()
    //   2. add test when manually create new StylesTable()
}
