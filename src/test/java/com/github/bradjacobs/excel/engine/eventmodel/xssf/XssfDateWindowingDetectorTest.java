package com.github.bradjacobs.excel.engine.eventmodel.xssf;

import com.github.bradjacobs.excel.testutils.TestResourceUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

class XssfDateWindowingDetectorTest {

    @Test
    void is1904DateWindowingTrue() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("date1904.xlsx");

        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);
            XssfDateWindowingDetector detector = new XssfDateWindowingDetector();
            boolean is1904DateWindowing = detector.is1904DateWindowing(reader);
            assertTrue(is1904DateWindowing);
        }
    }

    @Test
    void is1904DateWindowingFalse() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("test_data.xlsx");

        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFReader reader = new XSSFReader(pkg);
            XssfDateWindowingDetector detector = new XssfDateWindowingDetector();
            boolean is1904DateWindowing = detector.is1904DateWindowing(reader);
            assertFalse(is1904DateWindowing);
        }
    }
}
