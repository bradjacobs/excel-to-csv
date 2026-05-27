/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

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
import org.apache.poi.xssf.binary.XSSFBCommentsTable;
import org.apache.poi.xssf.binary.XSSFBSharedStringsTable;
import org.apache.poi.xssf.binary.XSSFBSheetHandler;
import org.apache.poi.xssf.binary.XSSFBStylesTable;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

/*
TODO: NOTE this class references POI classes that are marked as "internal"
 */
public class XssfbEventSheetReader implements EventSheetReader {

    private static final XssfbDateWindowingDetector DATE_WINDOWING_DETECTOR
            = new XssfbDateWindowingDetector();
    private static final XSSFBCommentsTable EMPTY_COMMENTS = null;
    private static final boolean FORMULAS_NOT_RESULTS = false;
    private static final PoiSheetStreamProvider sheetStreamProvider = new PoiSheetStreamProvider();

    private final XSSFBReader reader;
    private final SheetConfig sheetConfig;
    private final SharedStrings sharedStrings;
    private final XSSFBStylesTable styles;
    private final DataFormatter dataFormatter;

    public static XssfbEventSheetReader create(
            OPCPackage pkg,
            SheetConfig config) {
       try {
            XSSFBReader reader = new XSSFBReader(pkg);
            XSSFBSharedStringsTable sharedStrings = new XSSFBSharedStringsTable(pkg);
            XSSFBStylesTable styles = reader.getXSSFBStylesTable();

            boolean requires1904DateWindowing =
                    DATE_WINDOWING_DETECTOR.is1904DateWindowing(reader);

            DataFormatter formatter =
                    new DateWindowingDataFormatter(
                            requires1904DateWindowing);

            return new XssfbEventSheetReader(
                    reader,
                    config,
                    sharedStrings,
                    styles,
                    formatter);
       }
       catch (IOException | SAXException | OpenXML4JException  e) {
            throw new IllegalStateException(
                    "Failed to initialize XMLSheetStreamParser: " + e.getMessage(), e);
       }
    }

    private XssfbEventSheetReader(
            XSSFBReader reader,
            SheetConfig config,
            SharedStrings sharedStrings,
            XSSFBStylesTable styles,
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
        XSSFBSheetHandler sheetHandler = createSheetHandler(inputStream, stringRowConsumer);
        sheetHandler.parse();
        return stringRowConsumer.generateRowDataList();
    }

    private StringRowConsumer createStringRowConsumer() {
        return StringRowConsumer.of(
                sheetConfig.skipBlankRows(),
                sheetConfig.skipBlankColumns()
        );
    }

    private XSSFBSheetHandler createSheetHandler(
            InputStream inputStream,
            StringRowConsumer stringRowConsumer) {
        return sheetConfig.skipHiddenCells()
                ? createVisibleAwareHandler(inputStream, stringRowConsumer)
                : createDefaultHandler(inputStream, stringRowConsumer);
    }

    private XSSFBSheetHandler createDefaultHandler(
            InputStream inputStream,
            StringRowConsumer stringRowConsumer) {

        SheetContentHandler handler = new SheetContentHandler(
                sheetConfig,
                stringRowConsumer
        );

        XssfbCustomSheetContentsHandlerWrapper handlerWrapper =
                new XssfbCustomSheetContentsHandlerWrapper(handler, dataFormatter);

        return new XSSFBSheetHandler(
                inputStream,
                styles,
                EMPTY_COMMENTS,
                sharedStrings,
                handlerWrapper,
                FORMULAS_NOT_RESULTS
        );
    }

    private XSSFBSheetHandler createVisibleAwareHandler(
            InputStream inputStream,
            StringRowConsumer stringRowConsumer) {

        SheetContext sheetContext = new SheetContext();
        SheetContentHandler handler = new VisibleAwareSheetContentHandler(
                sheetConfig,
                stringRowConsumer,
                sheetContext
        );

        XssfbCustomSheetContentsHandlerWrapper handlerWrapper =
                new XssfbCustomSheetContentsHandlerWrapper(handler, dataFormatter);

        return new XssfbVisibleAwareSheetHandler(
                inputStream,
                styles,
                EMPTY_COMMENTS,
                sharedStrings,
                handlerWrapper,
                FORMULAS_NOT_RESULTS,
                sheetContext
        );
    }
}
