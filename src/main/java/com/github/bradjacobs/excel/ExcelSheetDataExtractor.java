/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.file.Path;
import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.DASHES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.SPACES;

// todo javadocs
public interface ExcelSheetDataExtractor {
    
    // Variations of reading Excel Sheet via sheet Index.
    String[][] readExcelSheetData(File excelFile, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(Path excelFile, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(URL excelFileUrl, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(InputStream is, int sheetIndex) throws IOException;
    String[][] readExcelSheetData(File excelFile, int sheetIndex, String password) throws IOException;
    String[][] readExcelSheetData(Path excelFile, int sheetIndex, String password) throws IOException;
    String[][] readExcelSheetData(InputStream is, int sheetIndex, String password) throws IOException;

    // Variations of reading Excel Sheet via sheet name.
    String[][] readExcelSheetData(File excelFile, String sheetName) throws IOException;
    String[][] readExcelSheetData(Path excelFile, String sheetName) throws IOException;
    String[][] readExcelSheetData(URL excelFileUrl, String sheetName) throws IOException;
    String[][] readExcelSheetData(InputStream is, String sheetName) throws IOException;
    String[][] readExcelSheetData(File excelFile, String sheetName, String password) throws IOException;
    String[][] readExcelSheetData(Path excelFile, String sheetName, String password) throws IOException;
    String[][] readExcelSheetData(InputStream is, String sheetName, String password) throws IOException;
}
