/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.ExcelNumberFormat;
import org.apache.poi.ss.usermodel.FormulaError;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.usermodel.XSSFComment;

/**
 * IMPORTANT NOTE: this is a copy of the private XSSFBSheetContentsHandlerWrapper class.
 * The difference here is the error method has a different implementation
 *  to be consistent with the error behavior of the other implementations.
 * @see XSSFSheetXMLHandler
 */
/*
TODO: NOTE this class references POI classes that are marked as "internal"
 */
class XssfbCustomSheetContentsHandlerWrapper implements XSSFBSheetHandler.XSSFBSheetContentsHandler {
    private final XSSFSheetXMLHandler.SheetContentsHandler delegate;
    private final DataFormatter dataFormatter;

    /**
     * Creates a wrapper that forwards events to the XML sheet handler while formatting numeric
     * cells.
     *
     * @param delegate target handler compatible with the XML streaming API
     * @param dataFormatter formatter used for numeric and date cell rendering
     */
    public XssfbCustomSheetContentsHandlerWrapper(
            XSSFSheetXMLHandler.SheetContentsHandler delegate, DataFormatter dataFormatter) {
        this.delegate = delegate;
        this.dataFormatter = dataFormatter;
    }

    @Override
    public void startRow(int rowNum) {
        delegate.startRow(rowNum);
    }

    @Override
    public void endRow(int rowNum) {
        delegate.endRow(rowNum);
    }

    @Override
    public void stringCell(String cellReference, String value, XSSFComment comment) {
        delegate.cell(cellReference, value, comment);
    }

    @Override
    public void doubleCell(
            String cellReference, double value, XSSFComment comment, ExcelNumberFormat nf) {
        String formattedValue =
                dataFormatter.formatRawCellContents(value, nf.getIdx(), nf.getFormat());
        delegate.cell(cellReference, formattedValue, comment);
    }

    @Override
    public void booleanCell(String cellReference, boolean value, XSSFComment comment) {
        delegate.cell(cellReference, Boolean.toString(value), comment);
    }

    @Override
    public void errorCell(String cellReference, FormulaError fe, XSSFComment comment) {
        String errorValue = (fe != null) ? fe.getString() : "ERROR";
        delegate.cell(cellReference, errorValue, comment);
    }

    @Override
    public void headerFooter(String text, boolean isHeader, String tagName) {
        delegate.headerFooter(text, isHeader, tagName);
    }

    @Override
    public void endSheet() {
        delegate.endSheet();
    }
}
