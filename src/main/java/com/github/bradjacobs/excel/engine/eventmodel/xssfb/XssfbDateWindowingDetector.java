/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.XSSFBParseException;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;

import java.io.IOException;
import java.io.InputStream;

/*
TODO: NOTE this class references POI classes that are marked as "internal"
 */
class XssfbDateWindowingDetector {

    public boolean is1904DateWindowing(XSSFBReader reader) throws IOException, InvalidFormatException {
        try (InputStream workbookData = reader.getWorkbookData()) {
            WorkbookPropsHandler workbookPropsHandler = new WorkbookPropsHandler(workbookData);
            workbookPropsHandler.parse();
            return workbookPropsHandler.isUses1904DateWindowing();
        }
    }

    private static class WorkbookPropsHandler extends XSSFBSheetHandler {
        // the '153' was from comment in org.apache.poi.xssf.binary.XSSFBRecordType
        // "BrtWbProp(153), //Workbook prop contains 1904/1900-date based bit"
        private static final int EXPECTED_RECORD_TYPE_ID = 153;
        private boolean uses1904DateWindowing = false;

        public WorkbookPropsHandler(InputStream is) {
            super(is, null, null, null, null, false);
        }

        public boolean isUses1904DateWindowing() {
            return uses1904DateWindowing;
        }

        @Override
        public void handleRecord(int id, byte[] data) throws XSSFBParseException {
            if (id == EXPECTED_RECORD_TYPE_ID) {
                int flags = LittleEndian.getUShort(data);
                // bit 0 = date1904
                this.uses1904DateWindowing = (flags & 0x0001) != 0;
            }
        }
    }
}
