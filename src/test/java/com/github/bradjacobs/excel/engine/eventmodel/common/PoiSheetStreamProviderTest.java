/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.common;

import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

class PoiSheetStreamProviderTest {

    /**
     * Simple case of reading sheet data (sheet name, index, and stream)
     *  from an Excel file.
     */
    @Test
    public void readSheetData() throws OpenXML4JException, IOException {
        Path testFilePath = TestResourceUtil.getResourceFilePath("repro.xlsx");
        String[] expectedSheetNames = {
                "BadRow", "WithUnicode", "ExtraBlankRowsAfterData",
                "WithBlankColumns1", "WithBlankColumns2"
        };

        PoiSheetStreamProvider poiSheetStreamProvider = new PoiSheetStreamProvider();
        List<EventSheet> eventSheets = null;

        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);
            eventSheets = poiSheetStreamProvider.getSheets(reader);
            for (int i = 0; i < eventSheets.size(); i++) {
                EventSheet sheet = eventSheets.get(i);
                assertEquals(i, sheet.getIndex());
                assertEquals(expectedSheetNames[i], sheet.getName());
                assertNotNull(sheet.getInputStream());
            }
        }
        finally {
            if (eventSheets != null) {
                for (EventSheet eventSheet : eventSheets) {
                    IOUtils.closeQuietly(eventSheet.getInputStream());
                }
            }
        }
    }

    @Test
    public void testNullReaderParameter() {
        Exception thrown = assertThrows(IllegalArgumentException.class, () -> {
            PoiSheetStreamProvider poiSheetStreamProvider = new PoiSheetStreamProvider();
            poiSheetStreamProvider.getSheets(null);
        });
        assertEquals("Must provide an XSSFReader reader", thrown.getMessage());

    }
}
