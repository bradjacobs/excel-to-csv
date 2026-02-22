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

// TODO - this is an experiment of what it would look like
//   for a interface to handle permutations.
//   Maybe swap out for an alternate solution and scrap this in the future (tbd)
public interface ExcelSheetDataExtractor {

    InputStreamGenerator inputStreamGenerator = new InputStreamGenerator();

    // Variations of reading Excel Sheet via sheet Index.
    default String[][] readExcelSheetData(File excelFile, int sheetIndex) throws IOException {
        return readExcelSheetData(excelFile, sheetIndex, null);
    }
    default String[][] readExcelSheetData(Path excelFile, int sheetIndex) throws IOException {
        return readExcelSheetData(excelFile, sheetIndex, null);
    }
    default String[][] readExcelSheetData(URL excelFileUrl, int sheetIndex) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFileUrl), sheetIndex, null);
    }
    default String[][] readExcelSheetData(InputStream is, int sheetIndex) throws IOException {
        return readExcelSheetData(is, sheetIndex, null);
    }
    default String[][] readExcelSheetData(File excelFile, int sheetIndex, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetIndex, password);
    }
    default String[][] readExcelSheetData(Path excelFile, int sheetIndex, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetIndex, password);
    }
    String[][] readExcelSheetData(InputStream is, int sheetIndex, String password) throws IOException;

    // Variations of reading Excel Sheet via sheet name.
    default String[][] readExcelSheetData(File excelFile, String sheetName) throws IOException {
        return readExcelSheetData(excelFile, sheetName, null);
    }
    default String[][] readExcelSheetData(Path excelFile, String sheetName) throws IOException {
        return readExcelSheetData(excelFile, sheetName, null);
    }
    default String[][] readExcelSheetData(URL excelFileUrl, String sheetName) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFileUrl), sheetName, null);
    }
    default String[][] readExcelSheetData(InputStream is, String sheetName) throws IOException {
        return readExcelSheetData(is, sheetName, null);
    }
    default String[][] readExcelSheetData(File excelFile, String sheetName, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetName, password);
    }
    default String[][] readExcelSheetData(Path excelFile, String sheetName, String password) throws IOException {
        return readExcelSheetData(inputStreamGenerator.getInputStream(excelFile), sheetName, password);
    }
    String[][] readExcelSheetData(InputStream is, String sheetName, String password) throws IOException;


    // below is common code for sheet configuration builder
    abstract class AbstractSheetConfigBuilder<T, B extends AbstractSheetConfigBuilder<T, B>> {
        protected boolean autoTrim = true;
        protected boolean removeBlankRows = false;
        protected boolean removeBlankColumns = false;
        protected boolean removeInvisibleCells = false;
        protected Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract B self();

        /**
         * Whether to trim whitespace on cell values
         * @param autoTrim (defaults to true)
         */
        public B autoTrim(boolean autoTrim) {
            this.autoTrim = autoTrim;
            return self();
        }

        /**
         * Whether to remove any blank rows.
         * @param removeBlankRows (defaults to false)
         */
        public B removeBlankRows(boolean removeBlankRows) {
            this.removeBlankRows = removeBlankRows;
            return self();
        }

        /**
         * Whether to remove any blank columns.
         * @param removeBlankColumns (defaults to false)
         */
        public B removeBlankColumns(boolean removeBlankColumns) {
            this.removeBlankColumns = removeBlankColumns;
            return self();
        }

        /**
         * Whether to prune out invisible cells.
         *   invisible = cellHeight = 0 or cellWidth = 0
         * @param removeInvisibleCells (defaults to false)
         */
        public B removeInvisibleCells(boolean removeInvisibleCells) {
            this.removeInvisibleCells = removeInvisibleCells;
            return self();
        }

        public B sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeFlag(SPACES, sanitizeSpaces);
        }

        public B sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeFlag(QUOTES, sanitizeQuotes);
        }

        public B sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeFlag(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        public B sanitizeDashes(boolean sanitizeDashes) {
            return setSanitizeFlag(DASHES, sanitizeDashes);
        }

        private B setSanitizeFlag(SpecialCharacterSanitizer.CharSanitizeFlag flag, boolean shouldAdd) {
            if (shouldAdd) {
                this.charSanitizeFlags.add(flag);
            }
            else {
                this.charSanitizeFlags.remove(flag);
            }
            return self();
        }

        // Ability to set the entire Set of SpecialCharSanitizers has limited access.
        protected B charSanitizeFlags(Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags) {
            this.charSanitizeFlags = charSanitizeFlags;
            return self();
        }

        protected SheetConfig buildConfig() {
            return new SheetConfig(
                    removeBlankRows,
                    removeBlankColumns,
                    removeInvisibleCells,
                    autoTrim,
                    charSanitizeFlags);
        }

        abstract public T build();
    }
}
