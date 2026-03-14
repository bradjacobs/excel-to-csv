/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.util;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.nio.file.Files;
import java.nio.file.Path;

public class TestExcelFileSheetUtils {

    public static Sheet getFileSheet(String fileName, String sheetName) {
        InputStream inputStream = TestResourceUtil.readResourceFileInputStream(fileName);
        try (inputStream; Workbook wb = WorkbookFactory.create(inputStream)) {
            Sheet sheet = wb.getSheet(sheetName);
            if (sheet == null) {
                throw new IllegalArgumentException("Sheet was not found: " + sheetName);
            }
            return sheet;
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to read Sheet from Excel File: " + e.getMessage(), e);
        }
    }


    public static Path createSingleCellExcelFile(Path tempDir, String cellValue) throws IOException {
        return createExcelFile(tempDir, new String[][]{{cellValue}});
    }

    public static Path createExcelFile(Path tempDir, String[][] sheetData) throws IOException {
        Path testFile = tempDir.resolve("tmp_" + System.currentTimeMillis() + ".xlsx");
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("TestCellSheet");

            for (int i = 0; i < sheetData.length; i++) {
                String[] rowData = sheetData[i];
                Row row = sheet.createRow(i);
                for (int j = 0; j < rowData.length; j++) {
                    String cellValue = rowData[j];
                    Cell cell = row.createCell(j);
                    cell.setCellValue(cellValue);
                }
            }

            try (FileOutputStream out = new FileOutputStream(testFile.toFile())) {
                workbook.write(out);
            }
            catch (IOException e) {
                throw new UncheckedIOException("Unable to create Temp Excel File: " + e.getMessage(), e);
            }
        }
        return testFile;
    }

    /**
     * Create a test Excel file that contains 1 sheet with 1 cell value.
     *   NOTE: assume that test @TempDir will automatically handle file clean up.
     * @param tempDir temp directory where the excel file should live
     * @param cellValue value to set for cell
     * @return Excel Sheet
     */
    public static Sheet createSingleCellExcelSheet(Path tempDir, String cellValue) throws IOException {
        Path testFile = createSingleCellExcelFile(tempDir, cellValue);

        try (InputStream inputStream = Files.newInputStream(testFile);
             Workbook wb = WorkbookFactory.create(inputStream)) {
            return wb.getSheetAt(0);
        }
        catch (IOException e) {
            throw new UncheckedIOException("Unable to read Sheet from Excel File: " + e.getMessage(), e);
        }
    }

    public static Path createSpecialExcelFile(Path outFile, String[][] sheetData) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook()) {
            XSSFSheet sheet = workbook.createSheet("TestCellSheet");

            // visible row/col values
            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("VISIBLE_A1");
            r0.createCell(1).setCellValue("HIDDEN_COL_B1");

            // hidden row values
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("HIDDEN_ROW_A2");
            r1.createCell(1).setCellValue("HIDDEN_ROW_AND_COL_B2");

            // Mark row 2 (index 1) hidden
            r1.setZeroHeight(true);

            // Mark column B (index 1) hidden
            sheet.setColumnHidden(1, true);




            try (FileOutputStream out = new FileOutputStream(outFile.toFile())) {
                workbook.write(out);
            }
            catch (IOException e) {
                throw new UncheckedIOException("Unable to create Temp Excel File: " + e.getMessage(), e);
            }
        }
        return outFile;
    }



}
