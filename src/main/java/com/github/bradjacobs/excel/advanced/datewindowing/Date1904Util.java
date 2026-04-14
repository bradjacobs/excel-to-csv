/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced.datewindowing;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;

public class Date1904Util {

    /**
     * Determines if the Excel file isDate1904 enabled
     * @param reader reader
     * @return true if date1904
     */
    public static boolean is1904DateWindowing(XSSFReader reader) throws IOException, InvalidFormatException, ParserConfigurationException, SAXException {
        WorkbookPropsHandler workbookPropsHandler =
                new WorkbookPropsHandler();
        try (InputStream workbookData = reader.getWorkbookData()) {
            SAXParserFactory.newInstance().newSAXParser().parse(workbookData, workbookPropsHandler);
        }
        return workbookPropsHandler.uses1904DateWindowing();
    }
}
