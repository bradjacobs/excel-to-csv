/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.common;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.xml.sax.SAXException;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;

public interface EventSheetReader {

    List<EventSheet> getSheets() throws IOException, InvalidFormatException;

    List<List<String>> read(InputStream inputStream) throws ParserConfigurationException, SAXException, IOException;
}
