/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;

/**
 * Extends the XSSFSheetXMLHandler to track hidden rows and columns.
 */
class VisibleAwareXSSFSheetXMLHandler extends XSSFSheetXMLHandler {

    private static final String ELEMENT_ROW = "row";
    private static final String ELEMENT_COL = "col";

    private static final String ATTR_ROW_NUMBER = "r";
    private static final String ATTR_HEIGHT = "ht";
    private static final String ATTR_HIDDEN = "hidden";
    private static final String ATTR_MIN = "min";
    private static final String ATTR_MAX = "max";

    private final SheetContext sheetContext;

    public VisibleAwareXSSFSheetXMLHandler(
            Styles styles,
            SharedStrings sharedStrings,
            SheetContentsHandler sheetContentsHandler,
            DataFormatter dataFormatter,
            boolean formulasNotResults,
            SheetContext sheetContext) {
        super(styles, sharedStrings, sheetContentsHandler, dataFormatter, formulasNotResults);
        this.sheetContext = sheetContext;
    }

    @Override
    public void startElement(String uri, String localName, String name, Attributes attributes) throws SAXException {
        if (ELEMENT_ROW.equals(name)) {
            handleRowStart(attributes);
        }
        else if (ELEMENT_COL.equals(name)) {
            handleColStart(attributes);
        }
        super.startElement(uri, localName, name, attributes);
    }

    private void handleRowStart(Attributes attributes) {
        boolean isHiddenRow = isZeroHeight(attributes) || isHidden(attributes);
        if (!isHiddenRow) {
            return;
        }

        int rowNumber1Based = getIntAttribute(attributes, ATTR_ROW_NUMBER);
        sheetContext.addHiddenRow(rowNumber1Based - 1);
    }

    private void handleColStart(Attributes attributes) {
        boolean isHiddenColumn = isHidden(attributes);
        if (!isHiddenColumn) {
            return;
        }

        int min1Based = getIntAttribute(attributes, ATTR_MIN);
        int max1Based = getIntAttribute(attributes, ATTR_MAX);

        // Columns are 1-based in XML, convert to 0-based index
        for (int col1Based = min1Based; col1Based <= max1Based; col1Based++) {
            sheetContext.addHiddenColumn(col1Based - 1);
        }
    }

    private boolean isZeroHeight(Attributes attributes) {
        String heightAttr = attributes.getValue(ATTR_HEIGHT);
        return "0".equals(heightAttr);
    }

    private boolean isHidden(Attributes attributes) {
        String hiddenAttr = attributes.getValue(ATTR_HIDDEN);
        return "1".equals(hiddenAttr) || "true".equalsIgnoreCase(hiddenAttr);
    }

    private int getIntAttribute(Attributes attributes, String attrName) {
        String value = attributes.getValue(attrName);
        if (value == null) {
            throw new IllegalArgumentException("Missing required attribute: " + attrName);
        }
        return Integer.parseInt(value);
    }
}
