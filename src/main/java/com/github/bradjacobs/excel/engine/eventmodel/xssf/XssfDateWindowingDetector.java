/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.DefaultHandler;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;

class XssfDateWindowingDetector {

    public boolean is1904DateWindowing(XSSFReader reader) throws IOException, InvalidFormatException, ParserConfigurationException, SAXException {
        WorkbookPropsHandler workbookPropsHandler =
                new WorkbookPropsHandler();
        try (InputStream workbookData = reader.getWorkbookData()) {
            SAXParserFactory.newInstance().newSAXParser().parse(workbookData, workbookPropsHandler);
        }
        return workbookPropsHandler.uses1904DateWindowing();
    }

    /**
     * Used to grab the 1904DateWindowing value from the Excel workbook.
     */
    private static class WorkbookPropsHandler extends DefaultHandler {
        private static final String WORKBOOK_PROPERTIES_ELEMENT = "workbookPr";
        private static final String DATE_1904_ATTRIBUTE = "date1904";

        private boolean uses1904DateWindowing = false;

        public boolean uses1904DateWindowing() {
            return uses1904DateWindowing;
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes) {
            if (isWorkbookPropertiesElement(localName, qName)) {
                uses1904DateWindowing = isEnabled(attributes.getValue(DATE_1904_ATTRIBUTE));
            }
        }

        private boolean isWorkbookPropertiesElement(String localName, String qName) {
            return WORKBOOK_PROPERTIES_ELEMENT.equals(localName) ||
                    WORKBOOK_PROPERTIES_ELEMENT.equals(qName);
        }

        private boolean isEnabled(String value) {
            return "1".equals(value) || "true".equalsIgnoreCase(value);
        }
    }
}
