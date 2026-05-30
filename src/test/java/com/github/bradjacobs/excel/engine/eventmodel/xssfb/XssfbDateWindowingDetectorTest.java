package com.github.bradjacobs.excel.engine.eventmodel.xssfb;

import com.github.bradjacobs.excel.util.TestResourceUtil;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.eventusermodel.XSSFBReader;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertTrue;

public class XssfbDateWindowingDetectorTest {

    @Test
    public void is1904DateWindowingTrue() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("date1904.xlsb");

        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFBReader reader = new XSSFBReader(pkg);
            XssfbDateWindowingDetector detector = new XssfbDateWindowingDetector();
            boolean is1904DateWindowing = detector.is1904DateWindowing(reader);
            assertTrue(is1904DateWindowing);
        }
    }

    @Test
    public void is1904DateWindowingFalse() throws Exception {
        Path testFilePath = TestResourceUtil.getResourceFilePath("test_data.xlsb");

        try (OPCPackage pkg = OPCPackage.open(testFilePath.toFile())) {
            XSSFBReader reader = new XSSFBReader(pkg);
            XssfbDateWindowingDetector detector = new XssfbDateWindowingDetector();
            boolean is1904DateWindowing = detector.is1904DateWindowing(reader);
            assertFalse(is1904DateWindowing);
        }
    }
}
