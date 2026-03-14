/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.StringRowConsumer;
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

    @Test
    void create_whenRemoveInvisibleCellsEnabled_returnsVisibleAwareHandler() {
        SheetConfig config = new SheetConfig(
                false,   // removeBlankRows
                false,   // removeBlankColumns
                true,    // removeInvisibleCells
                true,    // autoTrim
                Collections.emptySet()
        );

        StringRowConsumer rowConsumer = StringRowConsumer.of(false, false);
        SharedStrings sharedStrings = mock(SharedStrings.class);
        StylesTable styles = mock(StylesTable.class);

        ContentHandler handler = SheetContentHandlerFactory.create(
                config,
                rowConsumer,
                sharedStrings,
                styles
        );

        assertNotNull(handler);
        assertInstanceOf(VisibleAwareXSSFSheetXMLHandler.class, handler);
    }

    @Test
    void create_whenRemoveInvisibleCellsDisabled_returnsDefaultXssfHandler() {
        SheetConfig config = new SheetConfig(
                false,   // removeBlankRows
                false,   // removeBlankColumns
                false,   // removeInvisibleCells
                true,    // autoTrim
                Collections.emptySet()
        );

        StringRowConsumer rowConsumer = StringRowConsumer.of(false, false);
        SharedStrings sharedStrings = mock(SharedStrings.class);
        StylesTable styles = mock(StylesTable.class);

        ContentHandler handler = SheetContentHandlerFactory.create(
                config,
                rowConsumer,
                sharedStrings,
                styles
        );

        assertNotNull(handler);
        assertInstanceOf(XSSFSheetXMLHandler.class, handler);
    }
}