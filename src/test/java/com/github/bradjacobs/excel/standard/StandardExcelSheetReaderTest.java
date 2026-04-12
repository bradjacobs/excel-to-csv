/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.standard;

import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReaderTest;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.TestInstance;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.CsvSource;

import java.net.URL;
import java.nio.file.Path;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileUrl;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class StandardExcelSheetReaderTest extends AbstractExcelSheetReaderTest<StandardExcelSheetReader, StandardExcelSheetReader.Builder> {

    @Override
    protected StandardExcelSheetReader.Builder createBuilder() {
        return StandardExcelSheetReader.builder();
    }

    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");


    // NOTE - at the moment only the standard reader handles password files.
    // Happy Path testcase reading an Excel file that is password-protected.
    @Test
    public void testReadPasswordProtectedFile() throws Exception {
        URL resourceUrl = getResourceFileUrl("test_data_w_pswd_1234.xlsx");

        StandardExcelSheetReader sheetReader = StandardExcelSheetReader.builder().build();
        ExcelSheetReadRequest req = ExcelSheetReadRequest
                .from(resourceUrl)
                .password("1234")
                .build();

        SheetContent sheetContent = sheetReader.readSheet(req);
        String[][] dataMatrix = sheetContent.getMatrix();
        assertEquals("aaa", dataMatrix[0][0]);
        assertEquals("bbb", dataMatrix[0][1]);
    }

    @Nested
    @DisplayName("Invalid Password Tests")
    @TestInstance(TestInstance.Lifecycle.PER_CLASS)
    class InvalidPasswordTests {
        @ParameterizedTest
        @CsvSource({
                "bad_password, Password incorrect",
                ",no password was supplied",  // first param is null password
                "'',no password was supplied" // first param is empty string password
        })
        public void invalidPasswordCheck(String password, String expectedMessageSubstring) {
            StandardExcelSheetReader sheetReader = StandardExcelSheetReader.builder().build();
            ExcelSheetReadRequest req = ExcelSheetReadRequest
                    .from(VALID_TEST_INPUT_PSWD_PATH)
                    .password(password)
                    .build();

            Exception exception = assertThrows(Exception.class, () -> {
                sheetReader.readSheet(req);
            });
            assertContains(expectedMessageSubstring, exception.getMessage());
        }

        private void assertContains(String subString, String mainString) {
            assertTrue(mainString.contains(subString),
                    String.format("Expected to find substring '%s' in string '%s'.", subString, mainString));
        }
    }
}
