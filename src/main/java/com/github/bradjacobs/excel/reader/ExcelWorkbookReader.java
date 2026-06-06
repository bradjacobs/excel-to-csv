/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.reader;

import com.github.bradjacobs.excel.model.SheetContent;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import org.apache.commons.lang3.Validate;

import java.io.IOException;
import java.util.List;

// todo javadocs
public interface ExcelWorkbookReader {

    List<SheetContent> readSheets(ExcelReadRequest request) throws IOException;

    default SheetContent readSheet(ExcelReadRequest request) throws IOException {
        List<SheetContent> sheets = readSheets(request);
        Validate.isTrue(sheets.size() == 1, "Expected exactly one sheet but found " + sheets.size());
        return sheets.get(0);
    }
}
