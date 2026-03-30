/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import com.github.bradjacobs.excel.SheetData;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.Strings;
import org.apache.commons.lang3.Validate;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;
import java.util.function.Predicate;

/**
 * Simple version of a CSV Writer
 * For more advanced CSV writing features, it's
 * better to use one of the existing public libraries, such as
 * - OpenCSV
 * - Apache Commons CSV
 * - FastCSV
 */
public class CsvWriter {
    private static final Set<String> ALLOWED_OUTPUT_FILE_EXTENSIONS = Set.of("csv", "txt", "");
    private static final String NEW_LINE = System.lineSeparator();
    private static final String BOM = "\uFEFF"; // byte order marker for files with unicode.

    private final boolean saveUnicodeFileWithBom;
    private final Predicate<String> quoteRule;

    private CsvWriter(Builder builder) {
        this.quoteRule = builder.quoteMode.createPredicate();
        this.saveUnicodeFileWithBom = builder.saveUnicodeFileWithBom;
    }

    /**
     * Write the CSV data string out to a file.
     * @param sheetData the data to write out.
     * @param outputPath destination file.
     */
    public void writeToFile(SheetData sheetData, Path outputPath) throws IOException {
        validateOutputFileParameter(outputPath);
        String csvContent = toCsv(sheetData);
        Files.writeString(outputPath,
                prepareCsvContentForWriting(csvContent),
                StandardCharsets.UTF_8);
    }

    public void writeToFile(SheetData sheetData, File outputFile) throws IOException {
        writeToFile(sheetData, outputFile != null ? outputFile.toPath() : null);
    }

    /**
     * Convert the 2-D value matrix into a single CSV string.
     *   (quotes will be applied on values as needed)
     * @param sheetData string data matrix
     * @return csv file string
     */
    public String toCsv(SheetData sheetData) {
        String[][] dataMatrix = sheetData != null ? sheetData.getMatrix() : null;
        if (isEmptyDataMatrix(dataMatrix)) {
            return "";
        }

        StringBuilder sb = new StringBuilder();
        int columnCount = dataMatrix[0].length;
        int lastColumnIndex = columnCount - 1;

        for (String[] rowData : dataMatrix) {
            if (sb.length() != 0) {
                sb.append(NEW_LINE);
            }
            for (int i = 0; i < columnCount; i++) {
                String cellValue = rowData[i];
                if (quoteRule.test(cellValue)) {
                    // must first escape double quotes
                    cellValue = escapeDoubleQuotes(cellValue);
                    sb.append('\"').append(cellValue).append('\"');
                }
                else {
                    sb.append(cellValue);
                }

                if (i != lastColumnIndex) {
                    sb.append(',');
                }
            }
        }
        return sb.toString();
    }

    private String escapeDoubleQuotes(String value) {
        if (value.contains("\"")) {
            return Strings.CS.replace(value, "\"", "\"\"");
        }
        return value;
    }

    /**
     * checks if data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0;
    }

    private String prepareCsvContentForWriting(String csvContent) {
        // prepend the 'bom' so that Unicode characters will render correctly
        if (saveUnicodeFileWithBom && containsUnicode(csvContent)) {
            return BOM + csvContent;
        }
        return csvContent;
    }

    private boolean containsUnicode(String input) {
        return input.chars().anyMatch(ch -> ch > 127);
    }

    /**
     * Some sanity checks on the outputFile parameter
     * @param outputFile the destination CSV output file.
     * @throws IllegalArgumentException if a problem was detected with the File object.
     */
    private void validateOutputFileParameter(Path outputFile) throws IllegalArgumentException {
        Validate.isTrue(outputFile != null, "Must supply outputFile location to save CSV data.");
        Validate.isTrue(!Files.isDirectory(outputFile), "The outputFile cannot be an existing directory.");

        // confirm output file has an allowed file extension
        String ext = FilenameUtils.getExtension(outputFile.toString());
        Validate.isTrue(ALLOWED_OUTPUT_FILE_EXTENSIONS.contains(ext.toLowerCase()),
                "Illegal outputFile extension '%s'.  Must be either 'csv', 'txt' or blank", ext);

        Path parentDirectory = outputFile.toAbsolutePath().normalize().getParent();
        Validate.isTrue(parentDirectory != null && Files.isDirectory(parentDirectory),
                "Attempted to save CSV output file in a non-existent directory: " + outputFile);
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private QuoteMode quoteMode = QuoteMode.NORMAL;
        private boolean saveUnicodeFileWithBom = true;

        public Builder quoteMode(QuoteMode quoteMode) {
            Validate.isTrue(quoteMode != null, "QuoteMode cannot be null.");
            this.quoteMode = quoteMode;
            return this;
        }

        /**
         * Use a BOM when writing output file if data contains 'Unicode characters'
         * @param saveUnicodeFileWithBom (defaults to true)
         */
        public Builder saveUnicodeFileWithBom(boolean saveUnicodeFileWithBom) {
            this.saveUnicodeFileWithBom = saveUnicodeFileWithBom;
            return this;
        }

        public CsvWriter build() {
            return new CsvWriter(this);
        }
    }
}
