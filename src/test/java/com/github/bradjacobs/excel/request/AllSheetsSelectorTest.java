/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import com.github.bradjacobs.excel.util.TestSheetInfoUtil;
import org.junit.jupiter.api.Test;

import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class AllSheetsSelectorTest {

    @Test
    public void allSheetsSelectorHappyPath() {
        AllSheetsSelector selector = new AllSheetsSelector();
        List<SheetInfo> inputSheets = List.of(
                TestSheetInfoUtil.sheet("a1", 0),
                TestSheetInfoUtil.sheet("a2", 1)
        );

        List<SheetInfo> outputSheets = selector.filterSheets(inputSheets);
        assertEquals(inputSheets, outputSheets, "Expected all sheets to be returned");
    }
}
