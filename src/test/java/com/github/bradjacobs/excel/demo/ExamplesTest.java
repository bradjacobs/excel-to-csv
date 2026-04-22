package com.github.bradjacobs.excel.demo;

import com.github.bradjacobs.excel.ExcelProcessor;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.csv.CsvWriter;
import com.github.bradjacobs.excel.csv.QuoteMode;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.util.TestResourceUtil;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertTrue;

public class ExamplesTest {

    @TempDir
    private Path tempDir;

    @Test
    public void readExcelWriteCsv() throws IOException {
        Path inputFile = TestResourceUtil.getResourceFilePath("test_data.xlsx");
        Path outputFile = tempDir.resolve("output.csv");

        convertExcelToCsvFile(inputFile, outputFile);
        assertTrue(Files.exists(outputFile), "expected csv file was NOT created");
    }

    private static void convertExcelToCsvFile(Path inputFile, Path outputFile) throws IOException {
        ExcelSheetReadRequest request = ExcelSheetReadRequest.from(inputFile).build();
        SheetContent sheetContent = ExcelProcessor.builder().build().readSheet(request);
        CsvWriter.writeToFile(outputFile, sheetContent);
    }

    private static void convertExcelToCsvFile2(Path inputFile, Path outputFile) throws IOException {
        // create request to read 2 sheets from inputFile
        ExcelSheetReadRequest request = ExcelSheetReadRequest
                .from(inputFile)
                .bySheetNames("MySheet1", "MySheet2")
                .build();
        // create processor that will skip blank rows
        ExcelProcessor excelProcessor = ExcelProcessor.builder()
                .skipBlankRows(true)
                .build();
        // read the data from the 2 sheets
        List<SheetContent> sheetContentList = excelProcessor.readSheets(request);
        // write out 2 CSV files to given directory.  Only quote values required for CSV compliance.
        CsvWriter.writeToDirectory(Paths.get("outDir"), sheetContentList, QuoteMode.MINIMAL);
    }
}
