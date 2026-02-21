/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.ExcelSheetDataExtractor.AbstractSheetConfigBuilder;
import org.apache.commons.io.FilenameUtils;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Set;

/**
 * Simple class that reads an Excel Worksheet
 *   and will produce a CSV-equivalent
 */
// todo: add more method javadocs
public class ExcelReader {
    private final int sheetIndex;
    private final String sheetName;
    private final String password; // 'null' == no password
    private final boolean saveUnicodeFileWithBom;
    private final MatrixToCsvTextConverter matrixToCsvTextConverter;
    private final ExcelSheetReader excelSheetReader;

    private static final Set<String> ALLOWED_OUTPUT_FILE_EXTENSIONS = Set.of("csv", "txt", "");
    private static final String BOM = "\uFEFF"; // byte order marker for files with unicode.

    private final InputStreamGenerator inputStreamGenerator;

    private ExcelReader(Builder builder) {
        this.sheetIndex = builder.sheetIndex;
        this.sheetName = builder.sheetName;
        this.password = builder.password;
        this.saveUnicodeFileWithBom = builder.saveUnicodeFileWithBom;
        this.matrixToCsvTextConverter = new MatrixToCsvTextConverter(builder.quoteMode);
        this.excelSheetReader = createExcelSheetReader(builder);
        this.inputStreamGenerator = new InputStreamGenerator();
    }

    public void convertToCsvFile(Path excelFile, Path outputFile) throws IOException {
        validateOutputFileParameter(outputFile);
        writeCsvToFile( convertToCsvText(excelFile), outputFile);
    }
    public void convertToCsvFile(File excelFile, File outputFile) throws IOException {
        convertToCsvFile(fileToPath(excelFile), fileToPath(outputFile));
    }
    public void convertToCsvFile(URL excelUrl, Path outputFile) throws IOException {
        validateOutputFileParameter(outputFile);
        writeCsvToFile( convertToCsvText(excelUrl), outputFile);
    }
    public void convertToCsvFile(URL excelUrl, File outputFile) throws IOException {
        convertToCsvFile(excelUrl, fileToPath(outputFile));
    }

    public String convertToCsvText(Path excelFile) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelFile) );
    }
    public String convertToCsvText(File excelFile) throws IOException {
        return convertToCsvText(fileToPath(excelFile));
    }
    public String convertToCsvText(URL excelUrl) throws IOException {
        return matrixToCsvTextConverter.createCsvText( convertToDataMatrix(excelUrl) );
    }

    public String[][] convertToDataMatrix(Path excelFile) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelFile) );
    }
    public String[][] convertToDataMatrix(File excelFile) throws IOException {
        return convertToDataMatrix(fileToPath(excelFile));
    }
    public String[][] convertToDataMatrix(URL excelUrl) throws IOException {
        return convertToDataMatrix( inputStreamGenerator.getInputStream(excelUrl) );
    }


    private String[][] convertToDataMatrix(InputStream inputStream) throws IOException {
        if (StringUtils.isNotEmpty(this.sheetName)) {
            return excelSheetReader.readExcelSheetData(inputStream, this.sheetName, this.password);
        }
        else {
            return excelSheetReader.readExcelSheetData(inputStream, this.sheetIndex, this.password);
        }
    }

    /**
     * Write the CSV data string out to a file.
     * @param csvString CSV data
     * @param outputFile destination file.
     */
    private void writeCsvToFile(String csvString, Path outputFile) throws IOException {
        // prepend the 'bom' so that unicode characters will render correctly
        if (this.saveUnicodeFileWithBom && containsUnicode(csvString)) {
            csvString = BOM + csvString;
        }
        Files.writeString(outputFile, csvString, StandardCharsets.UTF_8);
    }

    /**
     * Some sanity checks on the outputFile parameter prior to doing the actual conversion
     *   (not exhaustive on all possible checks one could do)
     * @param outputFile the destination CSV output file.
     * @throws IllegalArgumentException if a problem was detected with the File object.
     */
    private void validateOutputFileParameter(Path outputFile) throws IllegalArgumentException {
        if (outputFile == null) {
            throw new IllegalArgumentException("Must supply outputFile location to save CSV data.");
        }
        else if (Files.isDirectory(outputFile)) {
            throw new IllegalArgumentException("The outputFile cannot be an existing directory.");
        }

        // confirm output file has an allowed file extension
        String ext = FilenameUtils.getExtension(outputFile.toString());
        if (! ALLOWED_OUTPUT_FILE_EXTENSIONS.contains(ext.toLowerCase())) {
            throw new IllegalArgumentException(
                    String.format("Illegal outputFile extension '%s'.  Must be either 'csv', 'txt' or blank", ext));
        }

        Path parentDirectory = outputFile.toAbsolutePath().normalize().getParent();
        if (parentDirectory == null || !Files.isDirectory(parentDirectory)) {
            throw new IllegalArgumentException("Attempted to save CSV output file in a non-existent directory: " + outputFile);
        }
    }

    private Path fileToPath(File file) {
        return (file != null ? file.toPath() : null);
    }

    private boolean containsUnicode(String input) {
        int inputLength = input.length();
        for (int i = 0; i < inputLength; i++) {
            int codePoint = input.codePointAt(i);
            if (codePoint > 127) {
                return true;
            }
        }
        return false;
    }

    /**
     * Creates an ExcelSheetReader based on the builder configuration.
     * @param builder builder
     * @return new ExcelSheetReader instance
     */
    protected ExcelSheetReader createExcelSheetReader(Builder builder) {
        return new ExcelSheetReader(builder.buildConfig());
    }

    public static Builder builder() {
        return new Builder();
    }

    // this builder extends abstract class to allow any of the
    //   ExcelSheetReader.Builder values to be set on this Builder as well.
    public static class Builder extends AbstractSheetConfigBuilder<ExcelReader, Builder> {
        private int sheetIndex = 0; // default to the first tab
        private String sheetName = ""; // optionally provide a specific sheet name
        private QuoteMode quoteMode = QuoteMode.NORMAL;
        private String password = null;
        private boolean saveUnicodeFileWithBom = true; // flag to write file with BOM if contains unicode.

        @Override
        protected Builder self() {
            return this;
        }

        /**
         * Set with sheet of Excel file to read (defaults to '0', i.e. the first sheet)
         * @param sheetIndex (0-based index of which sheet in Excel file to convert)
         */
        public Builder sheetIndex(int sheetIndex) {
            if (sheetIndex < 0) {
                throw new IllegalArgumentException("SheetIndex cannot be negative");
            }
            this.sheetIndex = sheetIndex;
            return this;
        }

        /**
         * Optionally can provide a sheet name instead of an index
         *  (if sheetName is set then sheetIndex is ignored)
         * @param sheetName name of Excel sheet
         */
        public Builder sheetName(String sheetName) {
            this.sheetName = sheetName;
            return this;
        }

        /**
         * Set how to handle quote/escaping string values to be CSV-compliant
         * @param quoteMode
         *  ALWAYS:  surround all values with quotes
         *  NORMAL:  add quotes around most values that contain non-alphanumeric (roughly similar to Jackson CsvMapper)
         *  LENIENT: add quotes around values that only really 'need' it to adhere to valid CSV (roughly similar to Excel 'save-as' CSV)
         *  NEVER:   never add quotes to any values.
         */
        public Builder quoteMode(QuoteMode quoteMode) {
            if (quoteMode == null) {
                throw new IllegalArgumentException("Cannot set quoteMode to null");
            }
            this.quoteMode = quoteMode;
            return this;
        }

        /**
         * Define a password to open the Excel file (if needed)
         * @param password excel fExcelReader.Buildeile password
         */
        public Builder password(String password) {
            // if user tries to set blank/empty string, then save as 'null'
            this.password = password != null && password.isEmpty() ? null : password;
            return this;
        }

        /**
         * Use a BOM when writing output file if data contains 'unicode characters'
         * @param saveUnicodeFileWithBom (defaults to true)
         */
        public Builder saveUnicodeFileWithBom(boolean saveUnicodeFileWithBom) {
            this.saveUnicodeFileWithBom = saveUnicodeFileWithBom;
            return this;
        }

        @Override
        public ExcelReader build() {
            return new ExcelReader(this);
        }
    }
}
