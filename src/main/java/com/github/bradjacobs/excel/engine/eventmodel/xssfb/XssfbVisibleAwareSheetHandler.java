/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import com.github.bradjacobs.excel.engine.eventmodel.common.SheetContext;
import org.apache.poi.util.LittleEndian;
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.binary.XSSFBParseException;
import org.apache.poi.xssf.binary.XSSFBRecordType;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.model.SharedStrings;

import java.io.InputStream;

class XssfbVisibleAwareSheetHandler extends XSSFBSheetHandler {

    private final SheetContext sheetContext;

    public XssfbVisibleAwareSheetHandler(
            InputStream is,
            XSSFBStylesTable styles,
            XSSFBCommentsTable comments,
            SharedStrings strings,
            XSSFBSheetContentsHandler sheetContentsHandler,
            boolean formulasNotResults,
            SheetContext sheetContext) {
        super(is, styles, comments, strings, sheetContentsHandler, formulasNotResults);
        this.sheetContext = sheetContext;
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
        // TODO - this criteria may be 'incomplete'
        //   'hidden' and 'zero height' could be different?
        int rowFlags = data[11] & 0xFF;

        boolean isHiddenRow = (rowFlags & 0x10) != 0;
        if (isHiddenRow) {
            int rowNumber = LittleEndian.getInt(data, 0);
            sheetContext.addHiddenRow(rowNumber);
        }
    }

    private void handleColumn(byte[] data) {
        // TODO - suspect data can be different lengths TBD
        int flags = data[16] & 0xFF;
        // Bit 0 = hidden
        boolean hidden = (flags & 0x0001) != 0;

        if (hidden) {
            // columns are marked with a 'range'
            int colFirst = LittleEndian.getInt(data, 0);
            int colLast  = LittleEndian.getInt(data, 4);

            for (int columnIndex = colFirst; columnIndex <= colLast; columnIndex++) {
                sheetContext.addHiddenColumn(columnIndex);
            }
        }
    }
}
