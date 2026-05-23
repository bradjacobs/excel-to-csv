/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.api.BasicSheetContent;
import com.github.bradjacobs.excel.api.SheetContent;
import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.core.AbstractExcelReader;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.request.SheetInfo;
import org.apache.commons.collections4.ListUtils;
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
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.extractor.ExtractorFactory.OOXML_PACKAGE;
import static org.apache.poi.poifs.crypt.Decryptor.DEFAULT_POIFS_ENTRY;

public class AdvancedExcelReader extends AbstractExcelReader {

    /**
     * Constructor
     */
    // TODO: Maybe change constructor to non-public
    public AdvancedExcelReader(SheetConfig config) {
        super(config);
    }

    @Override
    public List<SheetContent> readSheets(ExcelReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");

        InputStream sourceInputStream = request.getSourceInputStream();
        List<SheetContent> sheetContentList = new ArrayList<>();

        try (InputStream inputStream = preprocessFileInputStream(sourceInputStream, request.getPassword());
            OPCPackage pkg = OPCPackage.open(inputStream)) {
            XSSFReader reader = new XSSFReader(pkg);

            // NOTE: Docs say readOnlySharedStringsTable saves memory on large files,
            // but tests show ~25% slower performance!
            //reader.setUseReadOnlySharedStringsTable(true);

            List<SheetInfoRecord> allSheetInfos = fetchAllSheets(reader);

            try {
                List<SheetInfoRecord> selectedSheets =
                        request.getSheetSelector().filterSheets(allSheetInfos);
                List<SheetInfoRecord> unselectedSheets = ListUtils.subtract(allSheetInfos, selectedSheets);

                // close any 'extra' inputStreams that will not be processed.
                closeInputStreams(unselectedSheets);

                XMLSheetStreamReader xmlSheetStreamReader = XMLSheetStreamReader.create(sheetConfig, reader);

                for (SheetInfoRecord selectedSheet : selectedSheets) {
                    List<List<String>> sheetDataRows = xmlSheetStreamReader.read(selectedSheet.inputStream);
                    sheetContentList.add(new BasicSheetContent(selectedSheet.sheetName, sheetDataRows));
                }
            }
            finally {
                closeInputStreams(allSheetInfos);
            }
        }
        catch (OpenXML4JException | ParserConfigurationException | SAXException | IOException e) {
            throw new IOException("Failed to read Excel sheet data: " + e.getMessage(), e);
        }

        return sheetContentList;
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
                // if wrong type, then close the poifs and throw an exception below.
                poifs.close();
            }
        }

        if (FileMagic.OOXML != fm) {
            // todo - currently throw IOException to be consistent with the other impl.
            throw new IOException("Cannot open excel file - unsupported file type: " + fm);
            //throw new NotOfficeXmlFileException("Cannot open Excel file - unsupported file type: " + fm);
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
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractSheetConfigBuilder<AdvancedExcelReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public AdvancedExcelReader build() {
            return new AdvancedExcelReader(this.buildConfig());
        }
    }
}
