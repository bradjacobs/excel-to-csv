/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.shared;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

class DateWindowingDataFormatterTest {

    /**
     * Confirm the underlying `isEmulateCSV` is always set to 'true'
     * regardless of how DateWindowingDataFormatter was created.
     */
    @Test
    public void checkEmulateCsvFlag() {
        DataFormatter dataFormatterA = new DateWindowingDataFormatter(false);
        DataFormatter dataFormatterB = new DateWindowingDataFormatter(true);
        assertTrue(dataFormatterA.isEmulateCSV());
        assertTrue(dataFormatterB.isEmulateCSV());
    }

    @Test
    public void checkUseCachedValuesForFormulaCellsFlag() {
        DataFormatter dataFormatterA = new DateWindowingDataFormatter(false);
        DataFormatter dataFormatterB = new DateWindowingDataFormatter(true);
        assertTrue(dataFormatterA.useCachedValuesForFormulaCells());
        assertTrue(dataFormatterB.useCachedValuesForFormulaCells());
    }

    // todo add date tests
}
