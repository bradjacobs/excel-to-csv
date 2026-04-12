/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.csv;

import com.github.bradjacobs.excel.api.SheetContent;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.Validate;

import java.io.File;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
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
    private static final QuoteMode DEFAULT_MODE = QuoteMode.NORMAL;
    private static final boolean DEFAULT_SAVE_W_BOM = true;

    private final boolean saveUnicodeFileWithBom;
    private final Predicate<String> quoteRule;

    public static void write(Path outputPath, SheetContent sheetContent) throws IOException {
        write(outputPath, sheetContent, DEFAULT_MODE);
    }

    public static void write(Path outputPath, SheetContent sheetContent, QuoteMode quoteMode) throws IOException {
        CsvWriter csvWriter = new CsvWriter(quoteMode);
        csvWriter.writeToFile(outputPath, sheetContent);
    }

    public static void write(Path outputDirectory, List<SheetContent> sheetContents) throws IOException {
        write(outputDirectory, sheetContents, DEFAULT_MODE);
    }

    public static void write(Path outputDirectory, List<SheetContent> sheetContents, QuoteMode quoteMode) throws IOException {
        CsvWriter csvWriter = new CsvWriter(quoteMode);
        csvWriter.writeToDirectory(outputDirectory, sheetContents);
    }

    public CsvWriter() {
        this(DEFAULT_MODE, DEFAULT_SAVE_W_BOM);
    }
    public CsvWriter(QuoteMode quoteMode) {
        this(quoteMode, DEFAULT_SAVE_W_BOM);
    }
    public CsvWriter(QuoteMode quoteMode, boolean saveUnicodeFileWithBom) {
        Validate.isTrue(quoteMode != null, "QuoteMode cannot be null.");
        this.quoteRule = quoteMode.createPredicate();
        this.saveUnicodeFileWithBom = saveUnicodeFileWithBom;
    }

    /**
     * Write the CSV data string out to a file.
     * @param sheetContent the data to write out.
     * @param outputPath destination file.
     */
    public void writeToFile(Path outputPath, SheetContent sheetContent) throws IOException {
        validateOutputFileParameter(outputPath);
        String csvContent = toCsv(sheetContent);
        writeToFile(outputPath, csvContent);
    }

    public void writeToFile(File outputFile, SheetContent sheetContent) throws IOException {
        writeToFile(outputFile != null ? outputFile.toPath() : null, sheetContent);
    }

    public void writeToDirectory(Path outputDirectory, List<SheetContent> sheetContentList) throws IOException {
        Validate.isTrue(outputDirectory != null, "Must supply outputDirectory location to save CSV files.");
        Validate.isTrue(Files.isDirectory(outputDirectory), "Must supply a valid directory to write CSV data.");
        Validate.isTrue(sheetContentList != null && !sheetContentList.isEmpty(), "Must supply at least one sheetContent to write.");
        Validate.isTrue(!containsMissingSheetName(sheetContentList), "Must supply a non-empty sheetName for each sheetContent to write.");

        for (SheetContent sheetContent : sheetContentList) {
            String fileName = sheetContent.getSheetName().trim() + ".csv";
            Path outputPath = outputDirectory.resolve(fileName);
            String csvContent = toCsv(sheetContent);
            writeToFile(outputPath, csvContent);
        }
    }

    /**
     * Convert the 2-D value matrix into a single CSV string.
     *   (quotes will be applied on values as needed)
     * @param sheetContent string data matrix
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
                    sb.append(',');
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
     * checks if data matrix array is empty
     * @param dataMatrix dataMatrix
     * @return true if dataMatrix is considered 'empty'
     */
    private boolean isEmptyDataMatrix(String[][] dataMatrix) {
        return dataMatrix == null || dataMatrix.length == 0 || dataMatrix[0].length == 0;
    }

    private void writeToFile(Path outputPath, String csvContent) throws IOException {
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

        // confirm output file has an allowed file extension
        String ext = FilenameUtils.getExtension(outputFile.toString());
        Validate.isTrue(ALLOWED_OUTPUT_FILE_EXTENSIONS.contains(ext.toLowerCase()),
                "Illegal outputFile extension '%s'.  Must be either 'csv', 'txt' or blank", ext);

        Path parentDirectory = outputFile.toAbsolutePath().normalize().getParent();
        Validate.isTrue(parentDirectory != null && Files.isDirectory(parentDirectory),
                "Attempted to save CSV output file in a non-existent directory: " + outputFile);
    }

    private boolean containsMissingSheetName(List<SheetContent> sheetContentList) {
        return sheetContentList.stream()
                .anyMatch(obj -> obj == null ||
                        obj.getSheetName() == null ||
                        obj.getSheetName().trim().isEmpty());
    }
}
