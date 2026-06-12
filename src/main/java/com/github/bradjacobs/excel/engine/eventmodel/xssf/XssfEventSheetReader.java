/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.engine.eventmodel.shared.DateWindowingDataFormatter;
import com.github.bradjacobs.excel.engine.eventmodel.shared.EventSheet;
import com.github.bradjacobs.excel.engine.eventmodel.shared.EventSheetReader;
import com.github.bradjacobs.excel.engine.eventmodel.shared.PoiSheetStreamProvider;
import com.github.bradjacobs.excel.engine.eventmodel.shared.SheetContentHandler;
import com.github.bradjacobs.excel.engine.eventmodel.shared.SheetVisibilityTracker;
import com.github.bradjacobs.excel.engine.row.StringRowConsumer;
import com.github.bradjacobs.excel.sanitize.CellValueSanitizer;
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
    private static final PoiSheetStreamProvider SHEET_STREAM_PROVIDER = new PoiSheetStreamProvider();
    private static final boolean FORMULAS_NOT_RESULTS = false;
    private static final String INITIALIZATION_FAILURE_MESSAGE = "Failed to initialize XssfEventSheetReader";

    private final XSSFReader reader;
    private final SheetConfig sheetConfig;
    private final SharedStrings sharedStrings;
    private final StylesTable styles;
    private final DataFormatter dataFormatter;

    public static XssfEventSheetReader create(
            OPCPackage pkg,
            SheetConfig config) {
        try {
            XSSFReader xssfReader = new XSSFReader(pkg);

            // SIDE NOTE: Docs say readOnlySharedStringsTable saves memory
            // on large files, but experiments show ~25% slower performance!
            //xssfReader.setUseReadOnlySharedStringsTable(true);

            return new XssfEventSheetReader(
                    xssfReader,
                    config,
                    getSharedStrings(xssfReader),
                    getStylesTable(xssfReader),
                    createDataFormatter(xssfReader)
            );
        }
        catch (IOException | SAXException | OpenXML4JException | ParserConfigurationException e) {
            throw new IllegalStateException(INITIALIZATION_FAILURE_MESSAGE + ": " + e.getMessage(), e);
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
        return SHEET_STREAM_PROVIDER.getSheets(reader);
    }

    @Override
    public List<List<String>> read(InputStream inputStream)
            throws SAXException, IOException {
        StringRowConsumer rowConsumer = createStringRowConsumer();
        parseSheet(inputStream, rowConsumer);
        return rowConsumer.generateRowDataList();
    }

    private void parseSheet(InputStream inputStream, StringRowConsumer rowConsumer)
            throws SAXException, IOException {
        XMLReader xmlReader = createXmlReader(rowConsumer);
        xmlReader.parse(new InputSource(inputStream));
    }

    private StringRowConsumer createStringRowConsumer() {
        return StringRowConsumer.of(
                sheetConfig.skipBlankRows(),
                sheetConfig.skipBlankColumns()
        );
    }

    private XMLReader createXmlReader(StringRowConsumer rowConsumer)
            throws SAXException {
        XMLReader xmlReader;
        try {
            xmlReader = XMLHelper.newXMLReader();
        }
        catch (ParserConfigurationException e) {
            throw new SAXException("SAX parser configuration error: " + e.getMessage(), e);
        }
        xmlReader.setContentHandler(createContentHandler(rowConsumer));
        return xmlReader;
    }

    private ContentHandler createContentHandler(StringRowConsumer rowConsumer) {
        return sheetConfig.skipHiddenCells()
                ? createVisibleAwareHandler(rowConsumer)
                : createDefaultHandler(rowConsumer);
    }

    private ContentHandler createDefaultHandler(StringRowConsumer rowConsumer) {
        return new XSSFSheetXMLHandler(
                styles,
                sharedStrings,
                new SheetContentHandler(
                        rowConsumer,
                        createCellValueSanitizer(sheetConfig)
                ),
                dataFormatter,
                FORMULAS_NOT_RESULTS
        );
    }

    private ContentHandler createVisibleAwareHandler(StringRowConsumer rowConsumer) {
        SheetVisibilityTracker sheetVisibilityTracker = new SheetVisibilityTracker();

        return new XssfVisibleAwareSheetXmlHandler(
                styles,
                sharedStrings,
                new SheetContentHandler(
                        rowConsumer,
                        createCellValueSanitizer(sheetConfig),
                        sheetVisibilityTracker
                ),
                dataFormatter,
                FORMULAS_NOT_RESULTS,
                sheetVisibilityTracker
        );
    }

    private static DataFormatter createDataFormatter(XSSFReader reader)
            throws IOException, ParserConfigurationException, InvalidFormatException, SAXException {
        return new DateWindowingDataFormatter(requires1904DateWindowing(reader));
    }

    private static SharedStrings getSharedStrings(XSSFReader reader)
            throws IOException, InvalidFormatException {
        SharedStrings sharedStrings = reader.getSharedStringsTable();
        return sharedStrings != null ? sharedStrings : new SharedStringsTable();
    }

    private static StylesTable getStylesTable(XSSFReader reader)
            throws IOException, InvalidFormatException {
        StylesTable stylesTable = reader.getStylesTable();
        return stylesTable != null ? stylesTable : new StylesTable();
    }

    private static boolean requires1904DateWindowing(XSSFReader reader)
            throws IOException, ParserConfigurationException, InvalidFormatException, SAXException {
        return DATE_WINDOWING_DETECTOR.is1904DateWindowing(reader);
    }

    private static CellValueSanitizer createCellValueSanitizer(SheetConfig sheetConfig) {
        return new CellValueSanitizer(sheetConfig.trimStringValues(), sheetConfig.getCharSanitizeFlags());
    }
}
