/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import com.github.bradjacobs.excel.api.SheetContent;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.Validate;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Locale;
import java.util.Map;
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
    private static final Set<Character> VALID_DELIMITERS = Set.of(',', '\t', ';', ':', '|');
    private static final String NEW_LINE = System.lineSeparator();
    private static final String BOM = "\uFEFF"; // byte order marker for files with unicode.

    private static final QuoteMode DEFAULT_MODE = QuoteMode.NORMAL;
    private static final char DEFAULT_DELIMITER = ',';
    private static final boolean DEFAULT_SAVE_W_BOM = true;
    private static final boolean DEFAULT_ALLOW_OVERWRITE_FILE = false;

    private final Predicate<String> quoteRule;
    private final char delimiter;
    private final boolean saveUnicodeFileWithBom;
    private final boolean allowOverwriteFile;

    public static void writeToFile(Path outputPath, SheetContent sheetContent) throws IOException {
        writeToFile(outputPath, sheetContent, DEFAULT_MODE);
    }

    public static void writeToFile(Path outputPath, SheetContent sheetContent, QuoteMode quoteMode) throws IOException {
        CsvWriter csvWriter = CsvWriter.builder().quoteMode(quoteMode).build();
        csvWriter.saveToFile(outputPath, sheetContent);
    }

    public static void writeToDirectory(Path outputDirectory, List<SheetContent> sheetContents) throws IOException {
        writeToDirectory(outputDirectory, sheetContents, DEFAULT_MODE);
    }

    public static void writeToDirectory(Path outputDirectory, List<SheetContent> sheetContents, QuoteMode quoteMode) throws IOException {
        CsvWriter csvWriter = CsvWriter.builder().quoteMode(quoteMode).build();
        csvWriter.saveToDirectory(outputDirectory, sheetContents);
    }

    private CsvWriter(Builder builder) {
        this.quoteRule = builder.quoteMode.createPredicate(builder.delimiter);
        this.delimiter = builder.delimiter;
        this.allowOverwriteFile = builder.allowOverwriteFile;
        this.saveUnicodeFileWithBom = builder.saveUnicodeFileWithBom;
    }

    /**
     * Write the CSV data string out to a file.
     * @param sheetContent the data to write out.
     * @param outputPath destination file.
     */
    public void saveToFile(Path outputPath, SheetContent sheetContent) throws IOException {
        validateOutputFileParameter(outputPath);
        String csvContent = toCsv(sheetContent);
        saveToFile(outputPath, csvContent);
    }

    public void saveToFile(File outputFile, SheetContent sheetContent) throws IOException {
        saveToFile(fileToPath(outputFile), sheetContent);
    }

    public void saveToDirectory(Path outputDirectory, List<SheetContent> sheetContentList) throws IOException {
        Validate.isTrue(outputDirectory != null, "Must supply outputDirectory location to save CSV files.");
        Validate.isTrue(Files.isDirectory(outputDirectory), "Must supply a valid directory to write CSV data.");
        Validate.isTrue(!CollectionUtils.isEmpty(sheetContentList), "Must supply at least one sheetContent to write.");
        Validate.isTrue(!containsMissingSheetName(sheetContentList), "Must supply a non-empty sheetName for each sheetContent to write.");

        Map<Path, SheetContent> fileToSheetContentMap = new LinkedHashMap<>();
        // first create all the output file paths and validate.
        for (SheetContent sheetContent : sheetContentList) {
            String fileName = sheetContent.getSheetName().trim() + ".csv";
            Path outputPath = outputDirectory.resolve(fileName);
            validateOutputFileParameter(outputPath);
            fileToSheetContentMap.put(outputPath, sheetContent);
        }

        // do the file writes.
        for (Map.Entry<Path, SheetContent> entry : fileToSheetContentMap.entrySet()) {
            Path outputPath = entry.getKey();
            SheetContent sheetContent = entry.getValue();
            String csvContent = toCsv(sheetContent);
            saveToFile(outputPath, csvContent);
        }
    }

    private static Path fileToPath(File file) {
        return file != null ? file.toPath() : null;
    }

    /**
     * Convert the sheetContent into a single CSV string.
     *   (quotes will be applied on values as needed)
     * @param sheetContent sheetContent
     * @return csv file string
     */
    public String toCsv(SheetContent sheetContent) {
        String[][] dataMatrix = sheetContent != null ? sheetContent.getMatrix() : null;
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
                    sb.append(this.delimiter);
                }
            }
        }
        return sb.toString();
    }

    private String escapeDoubleQuotes(String value) {
        if (value.contains("\"")) {
            return value.replace("\"", "\"\"");
        }
        return value;
    }

    /**
     * checks if the data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0;
    }

    private void saveToFile(Path outputPath, String csvContent) throws IOException {
        Files.writeString(outputPath,
                prepareCsvContentForWriting(csvContent),
                StandardCharsets.UTF_8);
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

        Path fileNamePath = outputFile.getFileName();
        String fileName = fileNamePath != null ? fileNamePath.toString() : "";
        // confirm the output file has an allowed file extension
        String ext = FilenameUtils.getExtension(fileName);
        Validate.isTrue(ALLOWED_OUTPUT_FILE_EXTENSIONS.contains(ext.toLowerCase(Locale.ROOT)),
                "Illegal outputFile extension '%s'.  Must be either 'csv', 'txt' or blank", ext);

        Path parentDirectory = outputFile.toAbsolutePath().normalize().getParent();
        Validate.isTrue(parentDirectory != null && Files.isDirectory(parentDirectory),
                "Attempted to save CSV output file in a non-existent directory: " + outputFile);

        Validate.isTrue(allowOverwriteFile || !Files.exists(outputFile),
                "Attempted to overwrite an existing file: " + fileName);
    }

    private boolean containsMissingSheetName(List<SheetContent> sheetContentList) {
        return sheetContentList.stream()
                .anyMatch(obj -> obj == null ||
                        obj.getSheetName() == null ||
                        obj.getSheetName().trim().isEmpty());
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder {
        private QuoteMode quoteMode = DEFAULT_MODE;
        private char delimiter = DEFAULT_DELIMITER;
        private boolean saveUnicodeFileWithBom = DEFAULT_SAVE_W_BOM;
        private boolean allowOverwriteFile = DEFAULT_ALLOW_OVERWRITE_FILE;

        private Builder() {}

        public Builder quoteMode(QuoteMode quoteMode) {
            Validate.isTrue(quoteMode != null, "QuoteMode cannot be null.");
            this.quoteMode = quoteMode;
            return this;
        }

        public Builder delimiter(char delimiter) {
            Validate.isTrue(VALID_DELIMITERS.contains(delimiter), "Invalid delimiter: " + delimiter);
            this.delimiter = delimiter;
            return this;
        }

        /**
         * Use a BOM when writing an output file if data contains 'Unicode characters'
         * @param saveUnicodeFileWithBom (defaults to true)
         */
        public Builder saveUnicodeFileWithBom(boolean saveUnicodeFileWithBom) {
            this.saveUnicodeFileWithBom = saveUnicodeFileWithBom;
            return this;
        }

        public Builder allowOverwriteFile(boolean allowOverwriteFile) {
            this.allowOverwriteFile = allowOverwriteFile;
            return this;
        }

        public CsvWriter build() {
            return new CsvWriter(this);
        }
    }
}
