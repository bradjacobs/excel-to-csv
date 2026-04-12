/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import com.github.bradjacobs.excel.csv.QuoteMode;
import com.github.bradjacobs.excel.request.ExcelSheetReadRequest;
import org.apache.commons.io.FileUtils;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.Arguments;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.File;
import java.net.URL;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.List;
import java.util.stream.Stream;

import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileObject;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFilePath;
import static com.github.bradjacobs.excel.util.TestResourceUtil.getResourceFileUrl;
import static com.github.bradjacobs.excel.util.TestResourceUtil.readResourceFileText;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

// TODO - implement
class ExcelProcessorTest {

    @Test
    public void excelProcessorTest() {
        // TODO - add tests to check selection of standard vs advanced sheet readers
    }
}
