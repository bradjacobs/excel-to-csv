package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.CellValueReader;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.StringRowConsumer;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;

/**
 * Creates the appropriate sheet ContentHandler implementation based on SheetConfig.
 *
 * Keeping this logic in a separate class reduces branching and constructor complexity in SheetXMLReader.
 */
final class SheetContentHandlerFactory {

    private static final boolean FORMULAS_NOT_RESULTS = false;

    private SheetContentHandlerFactory() {
        // utility class
    }

    static ContentHandler create(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SharedStrings sharedStrings,
            StylesTable styles
    ) {
        if (sheetConfig.isRemoveInvisibleCells()) {
            SheetContext sheetContext = new SheetContext();
            return new VisibleAwareXSSFSheetXMLHandler(
                    styles,
                    sharedStrings,
                    new VisibleOnlySpecialSheetHandler(
                            sheetConfig,
                            stringRowConsumer,
                            sheetContext
                    ),
                    CellValueReader.getDataFormatter(),
                    FORMULAS_NOT_RESULTS,
                    sheetContext
            );
        }

        return new XSSFSheetXMLHandler(
                styles,
                sharedStrings,
                new SpecialSheetHandler(
                        sheetConfig,
                        stringRowConsumer
                ),
                CellValueReader.getDataFormatter(),
                FORMULAS_NOT_RESULTS
        );
    }
}