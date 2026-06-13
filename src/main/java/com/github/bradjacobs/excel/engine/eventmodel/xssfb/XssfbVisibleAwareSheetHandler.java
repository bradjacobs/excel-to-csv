/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import com.github.bradjacobs.excel.engine.eventmodel.shared.SheetVisibilityTracker;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.binary.XSSFBParseException;
import org.apache.poi.xssf.binary.XSSFBRecordType;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.model.SharedStrings;

import java.io.InputStream;

/*
TODO: NOTE this class references POI classes that are marked as "internal"
 */
class XssfbVisibleAwareSheetHandler extends XSSFBSheetHandler {

    private final SheetVisibilityTracker sheetVisibilityTracker;

    public XssfbVisibleAwareSheetHandler(
            InputStream is,
            XSSFBStylesTable styles,
            XSSFBCommentsTable comments,
            SharedStrings strings,
            XSSFBSheetContentsHandler sheetContentsHandler,
            boolean formulasNotResults,
            SheetVisibilityTracker sheetVisibilityTracker) {
        super(is, styles, comments, strings, sheetContentsHandler, formulasNotResults);
        this.sheetVisibilityTracker = sheetVisibilityTracker;
    }

    @Override
    public void handleRecord(int recordType, byte[] data) throws XSSFBParseException {
        if (recordType == XSSFBRecordType.BrtRowHdr.getId()) {
            handleRow(data);
        }
        else if (recordType == XSSFBRecordType.BrtColInfo.getId()) {
            handleColumn(data);
        }
        super.handleRecord(recordType, data);
    }

    private void handleRow(byte[] data) {
        int rowFlags = Byte.toUnsignedInt(data[11]);

        boolean isHiddenRow = (rowFlags & 0x10) != 0;
        if (isHiddenRow) {
            int rowNumber = LittleEndian.getInt(data, 0);
            sheetVisibilityTracker.addHiddenRow(rowNumber);
        }
    }

    private void handleColumn(byte[] data) {
        // TODO - suspect data can be different lengths TBD
        int flags = Byte.toUnsignedInt(data[16]);
        // Bit 0 = hidden
        boolean hidden = (flags & 0x0001) != 0;

        if (hidden) {
            // columns are marked with a 'range'
            int colFirst = LittleEndian.getInt(data, 0);
            int colLast  = LittleEndian.getInt(data, 4);

            for (int columnIndex = colFirst; columnIndex <= colLast; columnIndex++) {
                sheetVisibilityTracker.addHiddenColumn(columnIndex);
            }
        }
    }
}
