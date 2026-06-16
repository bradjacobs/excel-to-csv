/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import org.apache.commons.lang3.Validate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

public class EventSheetProvider {

    public List<EventSheet> getSheets(XSSFReader reader) throws IOException, InvalidFormatException {
        Validate.isTrue(reader != null, "Must provide an XSSFReader reader");

        List<EventSheet> sheets = new ArrayList<>();
        XSSFReader.SheetIterator sheetIterator = reader.getSheetIterator();
        int sheetIndex = 0;

        while (sheetIterator.hasNext()) {
            InputStream sheetStream = sheetIterator.next();
            String sheetName = sheetIterator.getSheetName();
            sheets.add(new EventSheet(sheetIndex, sheetName, sheetStream));
            sheetIndex++;
        }

        return sheets;
    }
}
