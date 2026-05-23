package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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

public class XMLSheetStreamReader {

    private static final DateWindowingDetector DATE_WINDOWING_DETECTOR = new DateWindowingDetector();
    private static final boolean FORMULAS_NOT_RESULTS = false;

    private final SheetConfig sheetConfig;
    private final SharedStrings sharedStrings;
    private final StylesTable styles;
    private final DataFormatter dataFormatter;

    public static XMLSheetStreamReader create(
            SheetConfig config,
            XSSFReader reader) {
        try {
            SharedStrings sharedStrings = getSharedStrings(reader);
            StylesTable styles = getStylesTable(reader);
            boolean requires1904DateWindowing =
                    requires1904DateWindowing(reader);

            DataFormatter formatter =
                    new DateWindowingDataFormatter(
                            requires1904DateWindowing);

            return new XMLSheetStreamReader(
                    config,
                    sharedStrings,
                    styles,
                    formatter);
        }
        catch (IOException | SAXException | InvalidFormatException | ParserConfigurationException e) {
            throw new IllegalStateException(
                    "Failed to initialize XMLSheetStreamParser: " + e.getMessage(), e);
        }
    }

    private XMLSheetStreamReader(
            SheetConfig config,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter) {
        this.sheetConfig = config;
        this.sharedStrings = sharedStrings;
        this.styles = styles;
        this.dataFormatter = dataFormatter;
    }

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
        return new VisibleAwareXSSFSheetXMLHandler(
                styles,
                sharedStrings,
                new VisibleCellsSheetContentHandler(
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
