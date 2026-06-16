/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.shared;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelDataFormattersTest {

    @Test
    void defaultFormatterEmulateCsvFlag() {
        DataFormatter dataFormatter = ExcelDataFormatters.standard();
        assertTrue(dataFormatter.isEmulateCSV());
    }

    @Test
    void defaultFormatterUseCachedValuesForFormulaCellsFlag() {
        DataFormatter dataFormatter = ExcelDataFormatters.standard();
        assertTrue(dataFormatter.useCachedValuesForFormulaCells());
    }

    /**
     * Confirm the underlying `isEmulateCSV` is always set to 'true'
     * regardless of how DateWindowingDataFormatter was created.
     */
    @Test
    void windowingFormatterEmulateCsvFlag() {
        DataFormatter dataFormatterA = ExcelDataFormatters.withDateWindowing(false);
        DataFormatter dataFormatterB = ExcelDataFormatters.withDateWindowing(true);
        assertTrue(dataFormatterA.isEmulateCSV());
        assertTrue(dataFormatterB.isEmulateCSV());
    }

    @Test
    void windowingFormatterUseCachedValuesForFormulaCellsFlag() {
        DataFormatter dataFormatterA = ExcelDataFormatters.withDateWindowing(false);
        DataFormatter dataFormatterB = ExcelDataFormatters.withDateWindowing(true);
        assertTrue(dataFormatterA.useCachedValuesForFormulaCells());
        assertTrue(dataFormatterB.useCachedValuesForFormulaCells());
    }

    // todo add some windowing date tests
}
