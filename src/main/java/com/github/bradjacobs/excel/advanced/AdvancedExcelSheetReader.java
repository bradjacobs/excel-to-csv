/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.advanced.datewindowing.WorkbookPropsHandler;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.NotOfficeXmlFileException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParserFactory;
import java.io.IOException;
import java.io.InputStream;

import static org.apache.poi.extractor.ExtractorFactory.OOXML_PACKAGE;
import static org.apache.poi.poifs.crypt.Decryptor.DEFAULT_POIFS_ENTRY;

public class AdvancedExcelSheetReader extends AbstractExcelSheetReader {

    private final SheetConfig sheetConfig;

    /**
     * Constructor
     */
    // TODO: Maybe change constructor to non-public
    public AdvancedExcelSheetReader(SheetConfig config) {
        this.sheetConfig = config;
    }

    @FunctionalInterface
    private interface SheetStreamProvider {
        InputStream fetch(XSSFReader reader) throws IOException, InvalidFormatException;
    }

    @FunctionalInterface
    private interface SheetMatcher {
        boolean isMatch(int sheetIndex, String sheetName);
    }

    @Override
    public String[][] readExcelSheetData(InputStream inputStream, int sheetIndex, String password) throws IOException {
        validateSheetIndex(sheetIndex);
        return readExcelSheetData(inputStream, password, reader -> fetchSheetInputStream(reader, sheetIndex));
    }

    @Override
    public String[][] readExcelSheetData(InputStream inputStream, String sheetName, String password) throws IOException {
        return readExcelSheetData(inputStream, password, reader -> fetchSheetInputStream(reader, sheetName));
    }

    private void validateSheetIndex(int sheetIndex) {
        if (sheetIndex < 0) {
            throw new IllegalArgumentException("Sheet Index cannot be negative");
        }
    }

    private String[][] readExcelSheetData(
            InputStream excelInputStream,
            String password,
            SheetStreamProvider sheetStreamProvider
    ) throws IOException {

        // preprocess the file stream to confirm it's the correct supported type,
        // and password decrypt the stream (if necessary)
        try (InputStream inputStream = preprocessFileInputStream(excelInputStream, password)) {
            // note: don't need to close the pkg in this context.
            OPCPackage pkg = OPCPackage.open(inputStream);
            XSSFReader reader = new XSSFReader(pkg);

            // NOTE - it's documented that the readOnlySharedStringsTable
            // setting can be more 'memory friendly' for very large files.
            // But preliminary testing shows setting readOnly to be ~ 25% SLOWER!
            //reader.setUseReadOnlySharedStringsTable(true);

            SheetXMLReader sheetXmlReader = createSheetXMLReader(reader);

            // get the inputStream for the specific sheet of interest
            try (InputStream sheetInputStream = sheetStreamProvider.fetch(reader)) {
                // actually parse the sheetInputStream for the data
                InputSource sheetSource = new InputSource(sheetInputStream);
                sheetXmlReader.parse(sheetSource);
                return sheetXmlReader.getSheetData();
            }
        }
        catch (OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IOException("Failed to read Excel sheet data: " + e.getMessage(), e);
        }
    }

    /**
     * Grab specific sheet inputStream based on sheetIndex.
     */
    private InputStream fetchSheetInputStream(
            XSSFReader reader,
            int sheetIndex
    ) throws IOException, InvalidFormatException {
        InputStream inputStream = fetchSheetInputStream(reader, (si, sn) -> sheetIndex == si);
        if (inputStream == null) {
            throw new IllegalArgumentException(String.format("Sheet index '%d' is out of range.", sheetIndex));
        }
        return inputStream;
    }

    /**
     * Grab specific sheet inputStream based on sheetName.
     */
    private InputStream fetchSheetInputStream(
            XSSFReader reader,
            String sheetName
    ) throws IOException, InvalidFormatException {
        InputStream inputStream = fetchSheetInputStream(reader, (si, sn) -> sn.equalsIgnoreCase(sheetName));
        if (inputStream == null) {
            throw new IllegalArgumentException(String.format("Sheet was not found: '%s'", sheetName));
        }
        return inputStream;
    }

    /**
     * Gets the specific sheet inputStream based on the selector.
     */
    private InputStream fetchSheetInputStream(
            XSSFReader reader,
            SheetMatcher sheetMatcher
    ) throws IOException, InvalidFormatException {

        // find the sheet inputStream that matches the selector,
        //  and close all the other sheet inputStreams.
        XSSFReader.SheetIterator sheetIterator = reader.getSheetIterator();
        int sheetIndex = 0;
        InputStream resultSheetInputStream = null;

        while (sheetIterator.hasNext()) {
            InputStream currentSheetStream = sheetIterator.next();
            String currentSheetName = sheetIterator.getSheetName();
            if (sheetMatcher.isMatch(sheetIndex, currentSheetName)) {
                resultSheetInputStream = currentSheetStream;
            }
            else {
                // close the stream if it's a non-matching sheet.
                currentSheetStream.close();
            }
            sheetIndex++;
        }

        return resultSheetInputStream;
    }

    private SheetXMLReader createSheetXMLReader(XSSFReader reader)
            throws IOException, InvalidFormatException, ParserConfigurationException, SAXException {

        SharedStrings sharedStrings = reader.getSharedStringsTable();
        StylesTable styles = reader.getStylesTable();

        WorkbookPropsHandler workbookPropsHandler =
                new WorkbookPropsHandler();
        try (InputStream workbookData = reader.getWorkbookData()) {
            SAXParserFactory.newInstance().newSAXParser().parse(workbookData, workbookPropsHandler);
        }
        boolean uses1904DateWindowing = workbookPropsHandler.uses1904DateWindowing();
        return new SheetXMLReader(this.sheetConfig, sharedStrings, styles, uses1904DateWindowing);
    }

    /**
     * Handle any input stream preprocessing. Such as decrypting
     * the stream if password protected and ensuring it's a supported OOXML type.
     */
    private InputStream preprocessFileInputStream(InputStream is, String password) throws IOException {
        InputStream resultStream = FileMagic.prepareToCheckMagic(is);
        FileMagic fm = FileMagic.valueOf(resultStream);

        if (FileMagic.OLE2 == fm) {
            POIFSFileSystem poifs = new POIFSFileSystem(resultStream);
            DirectoryNode root = poifs.getRoot();
            boolean isOOXML = root.hasEntryCaseInsensitive(DEFAULT_POIFS_ENTRY) || root.hasEntryCaseInsensitive(OOXML_PACKAGE);
            if (isOOXML) {
                fm = FileMagic.OOXML;
                // don't close root.getFileSystem() otherwise the container filesystem becomes invalid
                resultStream = DocumentFactoryHelper.getDecryptedStream(poifs, password);
            }
            else {
                // if wrong type then close the poifs and throw an exception below.
                poifs.close();
            }
        }

        if (FileMagic.OOXML != fm) {
            throw new NotOfficeXmlFileException("Cannot open excel file - unsupported file type: " + fm);
        }
        return resultStream;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractExcelSheetReader.AbstractSheetConfigBuilder<AdvancedExcelSheetReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public AdvancedExcelSheetReader build() {
            return new AdvancedExcelSheetReader(this.buildConfig());
        }
    }
}