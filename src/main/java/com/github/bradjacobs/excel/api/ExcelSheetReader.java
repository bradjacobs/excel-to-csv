/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import org.apache.commons.lang3.Validate;

import java.io.IOException;
import java.util.List;

// todo javadocs
public interface ExcelSheetReader {

    List<SheetContent> readSheets(ExcelSheetReadRequest request) throws IOException;

    default SheetContent readSheet(ExcelSheetReadRequest request) throws IOException {
        List<SheetContent> sheets = readSheets(request);
        Validate.isTrue(sheets.size() == 1, "Expected exactly one sheet but found " + sheets.size());
        return sheets.get(0);
    }
}
