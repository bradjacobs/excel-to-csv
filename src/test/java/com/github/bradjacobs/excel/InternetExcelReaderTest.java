/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.junit.jupiter.api.Disabled;
import org.junit.jupiter.api.Test;

import java.net.URL;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;

/**
 *   **** NOTE ****
 *   these tests are technically NOT 'unit tests' (which is why they are disabled)
 *   maybe later do the right thing and 'mock' the internet call if/when time allows.
 */
public class InternetExcelReaderTest {

    private static final String SAMPLE_INTERNET_EXCEL_FILE = "https://download.microsoft.com/download/1/4/E/14EDED28-6C58-4055-A65C-23B4DA81C4DE/Financial%20Sample.xlsx";

    @Disabled("Test is disabled because it makes external internet call")
    @Test
    public void testDownloadFromInternet() throws Exception {
        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.convertToCsvText(new URL(SAMPLE_INTERNET_EXCEL_FILE));
        assertNotNull(csvText);
        assertFalse(csvText.isEmpty());
    }
}
