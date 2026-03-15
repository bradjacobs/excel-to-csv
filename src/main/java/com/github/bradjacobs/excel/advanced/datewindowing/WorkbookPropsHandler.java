/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced.datewindowing;

import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

public class WorkbookPropsHandler extends DefaultHandler {

    private static final String WORKBOOK_PROPERTIES_ELEMENT = "workbookPr";
    private static final String DATE_1904_ATTRIBUTE = "date1904";

    private boolean uses1904DateWindowing;

    /**
     * Used to grab the 1904DateWindowing value from the Excel workbook.
     */
    @Override
    public void startElement(String uri, String localName, String qName, Attributes attrs) {
        boolean isWorkbookPropertiesElement =
                WORKBOOK_PROPERTIES_ELEMENT.equals(localName) ||
                        WORKBOOK_PROPERTIES_ELEMENT.equals(qName);

        if (isWorkbookPropertiesElement) {
            uses1904DateWindowing = isEnabled(attrs.getValue(DATE_1904_ATTRIBUTE));
        }
    }

    public boolean uses1904DateWindowing() {
        return uses1904DateWindowing;
    }

    private boolean isEnabled(String value) {
        return "1".equals(value) || "true".equalsIgnoreCase(value);
    }
}
