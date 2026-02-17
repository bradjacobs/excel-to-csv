package com.github.bradjacobs.excel;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
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
    abstract class AbstractSheetConfigBuilder<T extends AbstractSheetConfigBuilder<T>> {
        protected boolean autoTrim = true;
        protected boolean removeBlankRows = false;
        protected boolean removeBlankColumns = false;
        protected boolean removeInvisibleCells = false;
        protected Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags
                = new HashSet<>(SpecialCharacterSanitizer.DEFAULT_FLAGS);

        protected abstract T self();

        /**
         * Whether to trim whitespace on cell values
         * @param autoTrim (defaults to true)
         */
        public T autoTrim(boolean autoTrim) {
            this.autoTrim = autoTrim;
            return self();
        }

        /**
         * Whether to remove any blank rows.
         * @param removeBlankRows (defaults to false)
         */
        public T removeBlankRows(boolean removeBlankRows) {
            this.removeBlankRows = removeBlankRows;
            return self();
        }

        /**
         * Whether to remove any blank columns.
         * @param removeBlankColumns (defaults to false)
         */
        public T removeBlankColumns(boolean removeBlankColumns) {
            this.removeBlankColumns = removeBlankColumns;
            return self();
        }

        /**
         * Whether to prune out invisible cells.
         *   invisible = cellHeight = 0 or cellWidth = 0
         * @param removeInvisibleCells (defaults to false)
         */
        public T removeInvisibleCells(boolean removeInvisibleCells) {
            this.removeInvisibleCells = removeInvisibleCells;
            return self();
        }

        public T sanitizeSpaces(boolean sanitizeSpaces) {
            return setSanitizeFlag(SPACES, sanitizeSpaces);
        }

        public T sanitizeQuotes(boolean sanitizeQuotes) {
            return setSanitizeFlag(QUOTES, sanitizeQuotes);
        }

        public T sanitizeDiacritics(boolean sanitizeDiacritics) {
            return setSanitizeFlag(BASIC_DIACRITICS, sanitizeDiacritics);
        }

        public T sanitizeDashes(boolean sanitizeDashes) {
            return setSanitizeFlag(DASHES, sanitizeDashes);
        }

        private T setSanitizeFlag(SpecialCharacterSanitizer.CharSanitizeFlag flag, boolean shouldAdd) {
            if (shouldAdd) {
                this.charSanitizeFlags.add(flag);
            }
            else {
                this.charSanitizeFlags.remove(flag);
            }
            return self();
        }

        // Ability to set the entire Set of SpecialCharSanitizers has limited access.
        protected T charSanitizeFlags(Set<SpecialCharacterSanitizer.CharSanitizeFlag> charSanitizeFlags) {
            this.charSanitizeFlags =charSanitizeFlags;
            return self();
        }
    }
}
