/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.ContentHandler;

/**
 * Creates the appropriate sheet ContentHandler implementation based on SheetConfig.
 */
final class SheetContentHandlerFactory {

    private static final boolean FORMULAS_NOT_RESULTS = false;

    private SheetContentHandlerFactory() {}

    static ContentHandler create(
            SheetConfig sheetConfig,
            StringRowConsumer stringRowConsumer,
            SharedStrings sharedStrings,
            StylesTable styles,
            DataFormatter dataFormatter
    ) {
        if (sheetConfig.isRemoveInvisibleCells()) {
            SheetContext sheetContext = new SheetContext();
            return new VisibleAwareXSSFSheetXMLHandler(
                    styles,
                    sharedStrings,
                    new VisibleOnlySheetDataHandler(
                            sheetConfig,
                            stringRowConsumer,
                            sheetContext
                    ),
                    dataFormatter,
                    FORMULAS_NOT_RESULTS,
                    sheetContext
            );
        }

        return new XSSFSheetXMLHandler(
                styles,
                sharedStrings,
                new SheetDataHandler(
                        sheetConfig,
                        stringRowConsumer
                ),
                dataFormatter,
                FORMULAS_NOT_RESULTS
        );
    }
}