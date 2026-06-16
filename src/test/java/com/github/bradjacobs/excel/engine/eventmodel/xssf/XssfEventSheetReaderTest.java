/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import com.github.bradjacobs.excel.config.SheetConfig;
import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

class XssfEventSheetReaderTest {

    @Test
    void createThrowsIllegalStateExceptionWhenInitializationFails() throws Exception {
        // use the binary file as a way to test scenario where internal code throws an exception.
        Path binaryFilePath = TestResourceUtil.getResourceFilePath("test_data.xlsb");
        SheetConfig config = SheetConfig.builder().build();

        try (OPCPackage pkg = OPCPackage.open(binaryFilePath.toFile())) {
            IllegalStateException exception = assertThrows(
                    IllegalStateException.class,
                    () -> XssfEventSheetReader.create(pkg, config)
            );

            assertEquals(
                    "Failed to initialize XssfEventSheetReader: Invalid byte 1 of 1-byte UTF-8 sequence.",
                    exception.getMessage()
            );
        }
    }
}