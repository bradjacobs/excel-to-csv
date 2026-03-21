/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.api;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Path;

// todo javadocs
public interface ExcelSheetReader {

    // Read Excel first sheet.
    String[][] readExcelSheetData(File excelFile) throws IOException;
    String[][] readExcelSheetData(Path excelFile) throws IOException;
    String[][] readExcelSheetData(URL excelFileUrl) throws IOException;
    String[][] readExcelSheetData(InputStream inputStream) throws IOException;

    // Read Excel sheet by index.
    String[][] readExcelSheetData(File excelFile, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(Path excelFile, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(URL excelFileUrl, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(InputStream inputStream, int sheetIndex) throws IOException;

    // Read Excel sheet by index with password
    String[][] readExcelSheetData(File excelFile, int sheetIndex, String password) throws IOException;
    String[][] readExcelSheetData(Path excelFile, int sheetIndex, String password) throws IOException;
    String[][] readExcelSheetData(InputStream inputStream, int sheetIndex, String password) throws IOException;

    // Read Excel sheet by sheet name
    String[][] readExcelSheetData(File excelFile, String sheetName) throws IOException;
    String[][] readExcelSheetData(Path excelFile, String sheetName) throws IOException;
    String[][] readExcelSheetData(URL excelFileUrl, String sheetName) throws IOException;
    String[][] readExcelSheetData(InputStream inputStream, String sheetName) throws IOException;

    // Read Excel sheet by sheet name with password
    String[][] readExcelSheetData(File excelFile, String sheetName, String password) throws IOException;
    String[][] readExcelSheetData(Path excelFile, String sheetName, String password) throws IOException;
    String[][] readExcelSheetData(InputStream inputStream, String sheetName, String password) throws IOException;
}
