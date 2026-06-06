/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.engine.eventmodel.shared.EventSheet;
import com.github.bradjacobs.excel.engine.eventmodel.shared.EventSheetReader;
import com.github.bradjacobs.excel.engine.eventmodel.xssf.XssfEventSheetReader;
import com.github.bradjacobs.excel.engine.eventmodel.xssfb.XssfbEventSheetReader;
import com.github.bradjacobs.excel.model.BasicSheetContent;
import com.github.bradjacobs.excel.model.SheetContent;
import com.github.bradjacobs.excel.reader.AbstractExcelReader;
import com.github.bradjacobs.excel.request.ExcelReadRequest;
import com.github.bradjacobs.excel.request.SheetSelector;
import org.apache.commons.collections4.ListUtils;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.Validate;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.poifs.filesystem.DirectoryNode;
import org.apache.poi.poifs.filesystem.DocumentFactoryHelper;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFRelation;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.poi.extractor.ExtractorFactory.OOXML_PACKAGE;
import static org.apache.poi.poifs.crypt.Decryptor.DEFAULT_POIFS_ENTRY;

public class EventModelExcelReader extends AbstractExcelReader {

    /**
     * Constructor
     */
    // TODO: Maybe change constructor to non-public
    public EventModelExcelReader(SheetConfig config) {
        super(config);
    }

    @Override
    public List<SheetContent> readSheets(ExcelReadRequest request) throws IOException {
        Validate.isTrue(request != null, "Request cannot be null");

        InputStream sourceInputStream = request.getSourceInputStream();
        List<SheetContent> sheetContentList = new ArrayList<>();

        try (InputStream inputStream = preprocessFileInputStream(sourceInputStream, request.getPassword());
            OPCPackage pkg = OPCPackage.open(inputStream)) {

            EventSheetReader eventSheetReader =
                    createEventSheetReader(pkg, sheetConfig);

            List<EventSheet> selectedSheets = fetchSelectedSheets(eventSheetReader, request.getSheetSelector());

            try {
                for (EventSheet selectedSheet : selectedSheets) {
                    List<List<String>> sheetDataRows = eventSheetReader.read(selectedSheet.getInputStream());
                    sheetContentList.add(new BasicSheetContent(selectedSheet.getName(), sheetDataRows));
                }
            }
            finally {
                closeInputStreams(selectedSheets);
            }
        }
        catch (OpenXML4JException | ParserConfigurationException | SAXException | IOException e) {
            throw new IOException("Failed to read Excel sheet data: " + e.getMessage(), e);
        }

        return sheetContentList;
    }

    private void closeInputStreams(List<EventSheet> eventSheets) {
        for (EventSheet eventSheet : eventSheets) {
            IOUtils.closeQuietly(eventSheet.getInputStream());
        }
    }

    private List<EventSheet> fetchSelectedSheets(EventSheetReader eventSheetReader, SheetSelector sheetSelector) throws IOException, InvalidFormatException {
        List<EventSheet> allSheets = eventSheetReader.getSheets();
        List<EventSheet> selectedSheets = sheetSelector.filterSheets(allSheets);
        List<EventSheet> unselectedSheets = ListUtils.subtract(allSheets, selectedSheets);

        // close the inputStreams on the unselected sheets that will not be processed.
        closeInputStreams(unselectedSheets);
        return selectedSheets;
    }

    private EventSheetReader createEventSheetReader(
            OPCPackage pkg,
            SheetConfig sheetConfig) {

        if (isBinary(pkg)) {
            return XssfbEventSheetReader.create(pkg, sheetConfig);
        }
        return XssfEventSheetReader.create(pkg, sheetConfig);
    }

    private boolean isBinary(OPCPackage pkg) {
        POIXMLDocumentPart documentPart = new POIXMLDocumentPart(pkg);
        PackagePart packagePart = documentPart.getPackagePart();
        return XSSFRelation.XLSB_BINARY_WORKBOOK.getContentType()
                .equals(packagePart.getContentType());
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
            // NOTE: - currently throw IOException to be consistent with the
            // other standard implementations that throws from WorkbookFactory.create
            throw new IOException("Cannot open workbook - unsupported file type: " + fm);
        }
        return resultStream;
    }

    public static Builder builder() {
        return new Builder();
    }

    public static class Builder extends AbstractSheetConfigBuilder<EventModelExcelReader, Builder> {
        @Override
        protected Builder self() {
            return this;
        }

        @Override
        public EventModelExcelReader build() {
            return new EventModelExcelReader(this.buildConfig());
        }
    }
}
