/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced.datewindowing;

import org.apache.poi.xssf.eventusermodel.XSSFReader;

import javax.xml.parsers.SAXParserFactory;
import java.io.InputStream;

public class Date1904Util {

    /**
     * Determines if the Excel file isDate1904 enabled
     * @param reader reader
     * @return true if date1904
     */
    public static boolean is1904DateWindowing(XSSFReader reader) {
        WorkbookPropsHandler workbookPropsHandler =
                new WorkbookPropsHandler();
        try (InputStream workbookData = reader.getWorkbookData()) {
            SAXParserFactory.newInstance().newSAXParser().parse(workbookData, workbookPropsHandler);
        }
        catch (Exception e) {
            // todo should use better exception
            throw new RuntimeException("Error determining excel date1904 value: " + e.getMessage(), e);
        }

        return workbookPropsHandler.uses1904DateWindowing();
    }
}
