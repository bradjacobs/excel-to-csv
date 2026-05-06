/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.request;

import com.github.bradjacobs.excel.util.TestSheetInfoUtil;
import org.apache.commons.io.IOUtils;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelSheetReadRequestTest {

    @TempDir
    Path tempDir;

    private static final List<SheetInfo> TEST_SHEETS = List.of(
            TestSheetInfoUtil.sheet("AAA", 0),
            TestSheetInfoUtil.sheet("BBB", 1),
            TestSheetInfoUtil.sheet("CCC", 2)
    );

    private Path getSampleFilePath() {
        return tempDir.resolve("sample.xlsx");
    }


    @Nested
    @DisplayName("factory methods")
    class FactoryMethodsTests {
        @Test
        void forPathCreatesRequestWithDefaultSheetSelectorAndNoPassword() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath).build();

            assertEquals(filePath, request.getPath());
            assertNull(request.getUrl());
            assertNotNull(request.getSheetSelector());

            // confirm the 'default sheet selector' was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(1, filteredSheetList.size());
            assertEquals(0, filteredSheetList.get(0).getIndex());
            assertNull(request.getPassword());
        }

        @Test
        void forFileCreatesRequestUsingFilePath() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath.toFile()).build();

            assertEquals(filePath, request.getPath());
            assertNull(request.getUrl());
            assertNull(request.getPassword());
        }

        @Test
        void forPathWithNullCreatesRequestWithNullPathAndFailsOnBuild() {
            IllegalArgumentException exception = assertThrows(IllegalArgumentException.class,
                    () -> ExcelSheetReadRequest.from((Path)null).build());
            assertEquals("Either file path or url must be provided", exception.getMessage());
        }

        @Test
        void forFileWithNullCreatesRequestWithNullPathAndFailsOnBuild() {
            IllegalArgumentException exception = assertThrows(IllegalArgumentException.class,
                    () -> ExcelSheetReadRequest.from((File)null).build());
            assertEquals("Either file path or url must be provided", exception.getMessage());
        }

        @Test
        void forUrlCreatesRequestWithUrlSource() throws Exception {
            URL url = new URL("https://example.com/test.xlsx");
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(url).build();
            assertNull(request.getPath());
            assertEquals(url, request.getUrl());
            assertNull(request.getPassword());
        }
    }

    @Nested
    @DisplayName("optional configuration")
    class OptionalConfigurationTests {

        @Test
        void allowsOverridingDefaultSheetSelector() {
            Path filePath = getSampleFilePath();
            SheetSelector selector = new ByIndexSheetSelector(2);

            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .sheetSelector(selector)
                    .build();

            // confirm our custom sheet selector was applied.
            assertSame(selector, request.getSheetSelector());
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(1, filteredSheetList.size());
            assertEquals(2, filteredSheetList.get(0).getIndex());
        }

        @Test
        void selectAllSheets() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .allSheets()
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(TEST_SHEETS.size(), filteredSheetList.size());
            assertEquals(TEST_SHEETS, filteredSheetList);
        }

        @Test
        void byIndexSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetIndex(2)
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(1, filteredSheetList.size());
            assertEquals(2, filteredSheetList.get(0).getIndex());
        }

        @Test
        void byIndexVarArgsSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetIndexes(2, 1)
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(2, filteredSheetList.size());
            assertEquals(2, filteredSheetList.get(0).getIndex());
            assertEquals(1, filteredSheetList.get(1).getIndex());
        }

        @Test
        void byIndexCollectionSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetIndexes(List.of(2, 1))
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(2, filteredSheetList.size());
            assertEquals(2, filteredSheetList.get(0).getIndex());
            assertEquals(1, filteredSheetList.get(1).getIndex());
        }

        @Test
        void byNameSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetName("BBB")
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(1, filteredSheetList.size());
            assertEquals("BBB", filteredSheetList.get(0).getName());
        }

        @Test
        void byNameVarArgsSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetNames("CCC", "aaa")
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(2, filteredSheetList.size());
            assertEquals("CCC", filteredSheetList.get(0).getName());
            assertEquals("AAA", filteredSheetList.get(1).getName());
        }

        @Test
        void byNameCollectionSheetSelection() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .bySheetNames(List.of("CCC", "aaa"))
                    .build();

            // confirm our custom sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(2, filteredSheetList.size());
            assertEquals("CCC", filteredSheetList.get(0).getName());
            assertEquals("AAA", filteredSheetList.get(1).getName());
        }

        @Test
        void setsDefaultWithNullSheetSelector() {
            Path filePath = getSampleFilePath();

            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .sheetSelector(null)
                    .build();

            // confirm default sheet selector was applied.
            List<SheetInfo> filteredSheetList = request.getSheetSelector().filterSheets(TEST_SHEETS);
            assertEquals(1, filteredSheetList.size());
            assertEquals(0, filteredSheetList.get(0).getIndex());
        }

        @Test
        void allowsSettingPassword() {
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .password("secret")
                    .build();
            assertEquals("secret", request.getPassword());
        }

        @Test
        void setPasswordAsEmptyIsNull() {
            // setting password as empty string results in "no password".
            Path filePath = getSampleFilePath();
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(filePath)
                    .password("")
                    .build();
            assertEquals(null, request.getPassword());
        }
    }

    @Nested
    @DisplayName("source input stream")
    class SourceInputStreamTests {
        @Test
        void opensInputStreamFromPath() throws Exception {
            // NOTE: don't need a real Excel file for this testcase
            //  so just use a simple text file.
            String fileContent = "hello world";
            Path tempFile = createTempTextFile(fileContent);
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(tempFile).build();

            try (InputStream inputStream = request.getSourceInputStream()) {
                assertNotNull(inputStream);
                // check the content of the inputStream is expected.
                String content = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
                assertEquals(fileContent, content);
            }
        }

        @Test
        void opensFileInputStreamFromUrl() throws Exception {
            // NOTE: don't need a real Excel file for this testcase
            //  so just use a simple text file.
            String fileContent = "hello world";
            Path tempFile = createTempTextFile(fileContent);
            URL url = tempFile.toUri().toURL();

            // the request should recognized that the URL is a file, and read it.
            ExcelSheetReadRequest request = ExcelSheetReadRequest.from(url).build();
            try (InputStream inputStream = request.getSourceInputStream()) {
                assertNotNull(inputStream);
                // check the content of the inputStream is expected.
                String content = IOUtils.toString(inputStream, StandardCharsets.UTF_8);
                assertEquals(fileContent, content);
            }
        }
    }

    private Path createTempTextFile(String content) {
        Path tempFile = tempDir.resolve("test_"+System.currentTimeMillis()+".txt");
        try {
            Files.writeString(tempFile, content);
            return tempFile;
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to create temp file: " + e.getMessage(), e);
        }
    }
}