/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.advanced.datewindowing.DateWindowingDataFormatter;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;
import org.xml.sax.helpers.XMLFilterImpl;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;

class SheetXMLReader extends XMLFilterImpl {
    private final StringRowConsumer stringRowConsumer;

    public SheetXMLReader(
            SheetConfig sheetConfig,
            SharedStrings sharedStrings,
            StylesTable styles,
            boolean uses1904DateWindowing) throws ParserConfigurationException, SAXException {
        this.stringRowConsumer = StringRowConsumer.of(
                sheetConfig.skipBlankRows(),
                sheetConfig.skipBlankColumns()
        );
        DataFormatter dataFormatter = new DateWindowingDataFormatter(uses1904DateWindowing);
        XMLReader reader = createXmlReader(
                sheetConfig,
                sharedStrings,
                styles,
                dataFormatter);
        setParent(reader);
    }

    public void reset() {
        this.stringRowConsumer.reset();
    }

    // necessary override to correctly execute parse method.
    @Override
    public void parse(InputSource input) throws SAXException, IOException {
        getParent().parse(input);
    }

    public String[][] getSheetContentArray() {
        return this.stringRowConsumer.generateMatrix();
    }

    private XMLReader createXmlReader(
            SheetConfig sheetConfig,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter
    ) throws SAXException, ParserConfigurationException {
        XMLReader parser = XMLHelper.newXMLReader();
        parser.setContentHandler(
                SheetContentHandlerFactory.create(
                        sheetConfig,
                        stringRowConsumer,
                        sharedStrings,
                        styles,
                        dataFormatter
                )
        );
        return parser;
    }
}