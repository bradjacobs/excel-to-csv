/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertEquals;

class XssfDateWindowingDetectorTest {

    @Test
    void is1904DateWindowingTrue() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("date1904.xlsx");
        checkFile1904DateWindowing(testFilePath, true);
    }

    @Test
    void is1904DateWindowingFalse() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("test_data.xlsx");
        checkFile1904DateWindowing(testFilePath, false);
    }

    private void checkFile1904DateWindowing(Path testFilePath, boolean expected1904Windowing) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);
            XssfDateWindowingDetector detector = new XssfDateWindowingDetector();
            boolean is1904DateWindowing = detector.is1904DateWindowing(reader);
            assertEquals(expected1904Windowing, is1904DateWindowing);
        }
    }
}
