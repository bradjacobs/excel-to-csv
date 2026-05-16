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
import org.junit.jupiter.api.Test;
import org.xml.sax.ContentHandler;

import java.util.Collections;

import static org.junit.jupiter.api.Assertions.assertInstanceOf;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.mockito.Mockito.mock;

// TODO - this class was AI-generated
//   need to walk through and clean up as needed.
class SheetContentHandlerFactoryTest {

    private static final DataFormatter DATA_FORMATTER = new DataFormatter(true);

    @Test
    void create_whenSkipHiddenCellsEnabled_returnsVisibleAwareHandler() {
        SheetConfig config = new SheetConfig(
                false,   // skipBlankRows
                false,   // skipBlankColumns
                true,    // skipHiddenCells
                true,    // trimStringValues
                Collections.emptySet()
        );

        StringRowConsumer rowConsumer = StringRowConsumer.of(false, false);
        SharedStrings sharedStrings = mock(SharedStrings.class);
        StylesTable styles = mock(StylesTable.class);

        ContentHandler handler = SheetContentHandlerFactory.create(
                config,
                rowConsumer,
                sharedStrings,
                styles,
                DATA_FORMATTER
        );

        assertNotNull(handler);
        assertInstanceOf(VisibleAwareXSSFSheetXMLHandler.class, handler);
    }

    @Test
    void create_whenSkipInvisibleCellsDisabled_returnsDefaultXssfHandler() {
        SheetConfig config = new SheetConfig(
                false,   // skipBlankRows
                false,   // skipBlankColumns
                false,   // skipHiddenCells
                true,    // trimStringValues
                Collections.emptySet()
        );

        StringRowConsumer rowConsumer = StringRowConsumer.of(false, false);
        SharedStrings sharedStrings = mock(SharedStrings.class);
        StylesTable styles = mock(StylesTable.class);

        ContentHandler handler = SheetContentHandlerFactory.create(
                config,
                rowConsumer,
                sharedStrings,
                styles,
                DATA_FORMATTER
        );

        assertNotNull(handler);
        assertInstanceOf(XSSFSheetXMLHandler.class, handler);
    }
}