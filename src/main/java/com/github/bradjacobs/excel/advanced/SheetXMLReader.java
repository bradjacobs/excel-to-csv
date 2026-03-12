package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.StringRowConsumer;
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

    private final SheetConfig sheetConfig;
    private final StringRowConsumer stringRowConsumer;

    public SheetXMLReader(
            SheetConfig sheetConfig,
            SharedStrings sharedStrings,
            StylesTable styles) throws ParserConfigurationException, SAXException {
        this.sheetConfig = sheetConfig;
        this.stringRowConsumer = StringRowConsumer.of(
                sheetConfig.isRemoveBlankRows(),
                sheetConfig.isRemoveBlankColumns()
        );

        XMLReader reader = createSheetXmlReader(sharedStrings, styles);
        setParent(reader);
    }

    // necessary override to correctly execute parse method.
    @Override
    public void parse(InputSource input) throws SAXException, IOException {
        getParent().parse(input);
    }

    public String[][] getSheetData() {
        return this.stringRowConsumer.generateMatrix();
    }

    private XMLReader createSheetXmlReader(
            SharedStrings sharedStrings,
            StylesTable styles
    ) throws SAXException, ParserConfigurationException {
        XMLReader parser = XMLHelper.newXMLReader();
        parser.setContentHandler(
                SheetContentHandlerFactory.create(
                        sheetConfig,
                        stringRowConsumer,
                        sharedStrings,
                        styles
                )
        );
        return parser;
    }
}