/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel;

import com.github.bradjacobs.excel.engine.AbstractExcelReaderTest;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.apache.poi.util.XMLHelper;
import org.junit.jupiter.api.Test;
import org.mockito.MockedStatic;
import org.mockito.Mockito;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

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

    /**
     * Use mockito to check exception handling case
     * if 'XMLHelper.newXMLReader()' throws exception
     *   todo: try to find better home for this test
     */
    @Test
    void readThrowsRuntimeExceptionWhenNewXMLReaderFails() {
        Path filePath = TestResourceUtil.getResourceFilePath("test_data.xlsx");

        // Pass Mockito.CALLS_REAL_METHODS so all other XMLHelper methods behave normally
        try (MockedStatic<XMLHelper> mockedXmlHelper =
                     Mockito.mockStatic(XMLHelper.class, Mockito.CALLS_REAL_METHODS)) {
            String errorMessage = "Simulated Failure";
            // Only this specific method is mocked; everything else runs the real code
            mockedXmlHelper.when(XMLHelper::newXMLReader)
                    .thenThrow(new ParserConfigurationException(errorMessage));

            EventModelExcelReader reader = createBuilder().build();
            ExcelReadRequest request = ExcelReadRequest.from(filePath).build();

            Exception thrown = assertThrows(IOException.class, () -> {
                reader.readSheet(request);
            });
            assertTrue(thrown.getMessage().contains(errorMessage));
        }
    }

    // TODO
    //   1. add test when manually create new SharedStringsTable()
    //   2. add test when manually create new StylesTable()
}
