/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

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
import java.util.List;

class SheetXMLReader extends XMLFilterImpl {
    private final StringRowConsumer stringRowConsumer;

    public static SheetXMLReader create(
            SheetConfig sheetConfig,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter)
            throws ParserConfigurationException, SAXException {
        SheetXMLReader reader = new SheetXMLReader(sheetConfig);
        reader.init(sheetConfig, sharedStrings, styles, dataFormatter);
        return reader;
    }

    private SheetXMLReader(SheetConfig sheetConfig) {
        this.stringRowConsumer = StringRowConsumer.of(
                sheetConfig.skipBlankRows(),
                sheetConfig.skipBlankColumns()
        );
    }

    private void init(
            SheetConfig sheetConfig,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter)
            throws ParserConfigurationException, SAXException {

        XMLReader reader = createXmlReader(
                sheetConfig,
                this.stringRowConsumer,
                sharedStrings,
                styles,
                dataFormatter);
        setParent(reader);
    }

    // necessary override to correctly execute parse method.
    @Override
    public void parse(InputSource input) throws SAXException, IOException {
        getParent().parse(input);
    }

    public List<List<String>> getSheetDataRows() {
        return this.stringRowConsumer.generateMatrixList();
    }

    private XMLReader createXmlReader(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
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