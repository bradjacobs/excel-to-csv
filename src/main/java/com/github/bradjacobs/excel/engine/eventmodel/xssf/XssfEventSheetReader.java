/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import com.github.bradjacobs.excel.engine.eventmodel.common.DateWindowingDataFormatter;
import com.github.bradjacobs.excel.engine.eventmodel.common.EventSheet;
import com.github.bradjacobs.excel.engine.eventmodel.common.EventSheetReader;
import com.github.bradjacobs.excel.engine.eventmodel.common.PoiSheetStreamProvider;
import com.github.bradjacobs.excel.engine.eventmodel.common.SheetContentHandler;
import com.github.bradjacobs.excel.engine.eventmodel.common.SheetContext;
import com.github.bradjacobs.excel.engine.eventmodel.common.VisibleAwareSheetContentHandler;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.SharedStringsTable;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public class XssfEventSheetReader implements EventSheetReader {

    private static final XssfDateWindowingDetector DATE_WINDOWING_DETECTOR = new XssfDateWindowingDetector();
    private static final boolean FORMULAS_NOT_RESULTS = false;
    private static final PoiSheetStreamProvider sheetStreamProvider = new PoiSheetStreamProvider();

    private final XSSFReader reader;
    private final SheetConfig sheetConfig;
    private final SharedStrings sharedStrings;
    private final StylesTable styles;
    private final DataFormatter dataFormatter;

    public static XssfEventSheetReader create(
            OPCPackage pkg,
            SheetConfig config) {
        try {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sharedStrings = getSharedStrings(reader);
            StylesTable styles = getStylesTable(reader);
            boolean requires1904DateWindowing =
                    requires1904DateWindowing(reader);

            DataFormatter formatter =
                    new DateWindowingDataFormatter(
                            requires1904DateWindowing);

            return new XssfEventSheetReader(
                    reader,
                    config,
                    sharedStrings,
                    styles,
                    formatter);
        }
        catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException e) {
            throw new IllegalStateException(
                    "Failed to initialize XMLSheetStreamParser: " + e.getMessage(), e);
        }
    }

    private XssfEventSheetReader(
            XSSFReader reader,
            SheetConfig config,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter) {
        this.reader = reader;
        this.sheetConfig = config;
        this.sharedStrings = sharedStrings;
        this.styles = styles;
        this.dataFormatter = dataFormatter;
    }

    @Override
    public List<EventSheet> getSheets() throws IOException, InvalidFormatException {
        return sheetStreamProvider.getSheets(this.reader);
    }

    @Override
    public List<List<String>> read(InputStream inputStream) throws ParserConfigurationException, SAXException, IOException {
        StringRowConsumer stringRowConsumer = createStringRowConsumer();
        XMLReader xmlReader = createXmlReader(stringRowConsumer);
        xmlReader.parse(new InputSource(inputStream));
        return stringRowConsumer.generateRowDataList();
    }

    private StringRowConsumer createStringRowConsumer() {
        return StringRowConsumer.of(
                sheetConfig.skipBlankRows(),
                sheetConfig.skipBlankColumns()
        );
    }

    private XMLReader createXmlReader(
            StringRowConsumer stringRowConsumer
    ) throws SAXException, ParserConfigurationException {
        XMLReader parser = XMLHelper.newXMLReader();
        parser.setContentHandler(createSheetContentHandler(stringRowConsumer));
        return parser;
    }

    private ContentHandler createSheetContentHandler(StringRowConsumer stringRowConsumer) {
        return sheetConfig.skipHiddenCells()
                ? createVisibleAwareHandler(stringRowConsumer)
                : createDefaultHandler(stringRowConsumer);
    }

    private ContentHandler createVisibleAwareHandler(
            StringRowConsumer stringRowConsumer) {
        SheetContext sheetContext = new SheetContext();
        return new XssfVisibleAwareSheetXmlHandler(
                styles,
                sharedStrings,
                new VisibleAwareSheetContentHandler(
                        sheetConfig,
                        stringRowConsumer,
                        sheetContext
                ),
                dataFormatter,
                FORMULAS_NOT_RESULTS,
                sheetContext
        );
    }

    private ContentHandler createDefaultHandler(
            StringRowConsumer stringRowConsumer) {
        return new XSSFSheetXMLHandler(
                styles,
                sharedStrings,
                new SheetContentHandler(
                        sheetConfig,
                        stringRowConsumer
                ),
                dataFormatter,
                FORMULAS_NOT_RESULTS
        );
    }

    private static SharedStrings getSharedStrings(XSSFReader reader) throws IOException, InvalidFormatException {
        SharedStrings sharedStrings = reader.getSharedStringsTable();
        return sharedStrings != null ? sharedStrings : new SharedStringsTable();
    }

    private static StylesTable getStylesTable(XSSFReader reader) throws IOException, InvalidFormatException {
        StylesTable stylesTable = reader.getStylesTable();
        return stylesTable != null ? stylesTable : new StylesTable();
    }

    private static boolean requires1904DateWindowing(XSSFReader reader) throws IOException, ParserConfigurationException, InvalidFormatException, SAXException {
        return DATE_WINDOWING_DETECTOR.is1904DateWindowing(reader);
    }
}
