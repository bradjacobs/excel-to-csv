/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.advanced.datewindowing.Date1904Util;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelSheetReader;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import com.github.bradjacobs.excel.request.SheetInfo;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
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
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;

import static org.apache.poi.extractor.ExtractorFactory.OOXML_PACKAGE;
import static org.apache.poi.poifs.crypt.Decryptor.DEFAULT_POIFS_ENTRY;

public class AdvancedExcelSheetReader extends AbstractExcelSheetReader {

    /**
     * Constructor
     */
    // TODO: Maybe change constructor to non-public
    public AdvancedExcelSheetReader(SheetConfig config) {
        super(config);
    }

    @Override
    public List<SheetContent> readSheets(ExcelSheetReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");

        InputStream sourceInputStream = request.getSourceInputStream();
        if (sourceInputStream == null) {
            throw new IllegalArgumentException("Request must provide an InputStream");
        }

        List<SheetContent> sheetContentList = new ArrayList<>();

        try (InputStream inputStream = preprocessFileInputStream(sourceInputStream, request.getPassword())) {
            OPCPackage pkg = OPCPackage.open(inputStream);
            XSSFReader reader = new XSSFReader(pkg);
            SheetXMLReader sheetXmlReader = createSheetXMLReader(reader);
            List<SheetInfoRecord> allSheetInfos = fetchAllSheets(reader);

            try {
                List<SheetInfoRecord> selectedSheets =
                        request.getSheetSelector().filterSheets(allSheetInfos);

                closeInputStreams(getUnselectedSheets(allSheetInfos, selectedSheets));

                for (SheetInfoRecord selectedSheet : selectedSheets) {
                    sheetContentList.add(extractSheetContent(selectedSheet, sheetXmlReader));
                }
            }
            finally {
                closeInputStreams(allSheetInfos);
            }
        }
        catch (OpenXML4JException | ParserConfigurationException | SAXException e) {
            throw new IOException("Failed to read Excel sheet data: " + e.getMessage(), e);
        }

        return sheetContentList;
    }

    private List<SheetInfoRecord> getUnselectedSheets(List<SheetInfoRecord> allSheetInfos,
                                                      List<SheetInfoRecord> selectedSheets) {
        List<SheetInfoRecord> unselectedSheets = new ArrayList<>(allSheetInfos);
        unselectedSheets.removeAll(selectedSheets);
        return unselectedSheets;
    }

    private SheetContent extractSheetContent(SheetInfoRecord sheetInfoRecord, SheetXMLReader sheetXmlReader) throws IOException, SAXException, ParserConfigurationException {
        try (InputStream sheetInputStream = sheetInfoRecord.inputStream) {
            // actually parse the sheetInputStream for the data
            InputSource sheetSource = new InputSource(sheetInputStream);
            sheetXmlReader.parse(sheetSource);
            String[][] sheetValuesMatrix = sheetXmlReader.getSheetContentArray();

            // TODO - a bit kludgy.  Reset the read to can be used for the next sheet.
            sheetXmlReader.reset();
            return new SheetContent(sheetInfoRecord.sheetName, sheetValuesMatrix);
        }
    }

    private void closeInputStreams(List<SheetInfoRecord> sheetInfoRecords) {
        for (SheetInfoRecord sheetInfoRecord : sheetInfoRecords) {
            IOUtils.closeQuietly(sheetInfoRecord.inputStream);
        }
    }

    private List<SheetInfoRecord> fetchAllSheets(XSSFReader reader) throws IOException, InvalidFormatException {
        List<SheetInfoRecord> records = new ArrayList<>();
        XSSFReader.SheetIterator sheetIterator = reader.getSheetIterator();
        int sheetIndex = 0;

        while (sheetIterator.hasNext()) {
            InputStream sheetStream = sheetIterator.next();
            String sheetName = sheetIterator.getSheetName();
            records.add(new SheetInfoRecord(sheetIndex, sheetName, sheetStream));
            sheetIndex++;
        }
        return records;
    }

    private SheetXMLReader createSheetXMLReader(XSSFReader reader)
            throws IOException, InvalidFormatException, ParserConfigurationException, SAXException {

        SharedStrings sharedStrings = reader.getSharedStringsTable();
        StylesTable styles = reader.getStylesTable();
        boolean uses1904DateWindowing = Date1904Util.is1904DateWindowing(reader);

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
            // todo - currently throw IOException to be consistent with the other impl.
            throw new IOException("Cannot open excel file - unsupported file type: " + fm);
            //throw new NotOfficeXmlFileException("Cannot open excel file - unsupported file type: " + fm);
        }
        return resultStream;
    }

    // simple pojo to hold information for a specific sheet.
    private static class SheetInfoRecord implements SheetInfo {
        private final int sheetIndex;
        private final String sheetName;
        private final InputStream inputStream;

        public SheetInfoRecord(int sheetIndex, String sheetName, InputStream inputStream) {
            this.sheetIndex = sheetIndex;
            this.sheetName = sheetName;
            this.inputStream = inputStream;
        }

        @Override
        public String getName() {
            return sheetName;
        }

        @Override
        public int getIndex() {
            return sheetIndex;
        }

        @Override
        public boolean equals(Object o) {
            if (o == null || getClass() != o.getClass()) return false;
            SheetInfoRecord that = (SheetInfoRecord) o;
            return sheetIndex == that.sheetIndex;
        }

        @Override
        public int hashCode() {
            return Objects.hashCode(sheetIndex);
        }
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