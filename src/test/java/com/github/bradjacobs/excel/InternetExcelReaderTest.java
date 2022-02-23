package com.github.bradjacobs.excel;

import org.testng.annotations.Test;

import java.net.URL;

import static org.testng.Assert.assertNotNull;
import static org.testng.Assert.assertTrue;

/**
 *          **** NOTE ****
 *   these tests are technically NOT 'unit tests' (which is why they are disabled)
 *
 *   maybe later do the right thing and 'mock' the internet call if/when time allows.
 *
 */
public class InternetExcelReaderTest
{
    private static final boolean INTERNET_REQUESTS_ENABLED = false; // disable the tests!

    private static final String SAMPLE_INTERNET_EXCEL_FILE = "https://download.microsoft.com/download/1/4/E/14EDED28-6C58-4055-A65C-23B4DA81C4DE/Financial%20Sample.xlsx";

    @Test(enabled = INTERNET_REQUESTS_ENABLED)
    public void testDownloadFromInternet() throws Exception
    {
        ExcelReader excelReader = ExcelReader.builder().build();
        String csvText = excelReader.convertToCsvText(new URL(SAMPLE_INTERNET_EXCEL_FILE));
        assertNotNull(csvText);
        assertTrue(csvText.length() > 0);
    }
}
