package com.github.bradjacobs.excel.advanced;

import com.github.bradjacobs.excel.SheetConfig;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Set;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.*;

// TODO - this class was AI-generated
//   need to walk through and clean up as needed.
class SheetXMLReaderTest {

    @Test
    void parse_delegatesToParentAndProducesMatrix_basicVisibleSheet() throws Exception {
        byte[] workbookBytes = buildWorkbookBytes(wb -> {
            Sheet sheet = wb.createSheet("S1");
            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("A1");
            r0.createCell(1).setCellValue("B1");

            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("A2");
            r1.createCell(1).setCellValue("B2");
        });

        SheetConfig cfg = new SheetConfig(
                false, // removeBlankRows
                false, // removeBlankColumns
                false, // removeInvisibleCells
                true,  // autoTrim
                Set.of()
        );

        String[][] matrix = parseFirstSheet(workbookBytes, cfg);

        assertNotNull(matrix, "matrix must not be null");
        assertEquals(2, matrix.length, "expected two rows");
        assertArrayEquals(new String[]{"A1", "B1"}, matrix[0], "row 0 mismatch");
        assertArrayEquals(new String[]{"A2", "B2"}, matrix[1], "row 1 mismatch");
    }

    @Test
    void removeInvisibleCells_true_skipsHiddenRowAndHiddenColumn() throws Exception {
        byte[] workbookBytes = buildWorkbookBytes(wb -> {
            Sheet sheet = wb.createSheet("S1");

            // visible row/col values
            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("VISIBLE_A1");
            r0.createCell(1).setCellValue("HIDDEN_COL_B1");

            // hidden row values
            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("HIDDEN_ROW_A2");
            r1.createCell(1).setCellValue("HIDDEN_ROW_AND_COL_B2");

            // Mark row 2 (index 1) hidden
            r1.setZeroHeight(true);

            // Mark column B (index 1) hidden
            sheet.setColumnHidden(1, true);
        });

        SheetConfig cfg = new SheetConfig(
                false, // removeBlankRows
                false, // removeBlankColumns
                true,  // removeInvisibleCells
                true,  // autoTrim
                Set.of()
        );

        String[][] matrix = parseFirstSheet(workbookBytes, cfg);

        assertNotNull(matrix, "matrix must not be null");
        assertEquals(1, matrix.length, "hidden row should be excluded entirely");
        assertArrayEquals(new String[]{"VISIBLE_A1"}, matrix[0], "hidden column should be excluded");
    }

    @Test
    void removeInvisibleCells_false_includesHiddenRowAndHiddenColumn() throws Exception {
        byte[] workbookBytes = buildWorkbookBytes(wb -> {
            Sheet sheet = wb.createSheet("S1");

            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("VISIBLE_A1");
            r0.createCell(1).setCellValue("HIDDEN_COL_B1");

            Row r1 = sheet.createRow(1);
            r1.createCell(0).setCellValue("HIDDEN_ROW_A2");
            r1.createCell(1).setCellValue("HIDDEN_ROW_AND_COL_B2");

            r1.setZeroHeight(true);
            sheet.setColumnHidden(1, true);
        });

        SheetConfig cfg = new SheetConfig(
                false, // removeBlankRows
                false, // removeBlankColumns
                false, // removeInvisibleCells
                true,  // autoTrim
                Set.of()
        );

        String[][] matrix = parseFirstSheet(workbookBytes, cfg);

        assertNotNull(matrix, "matrix must not be null");
        assertEquals(2, matrix.length, "when not removing invisible cells, hidden row should remain");
        assertArrayEquals(new String[]{"VISIBLE_A1", "HIDDEN_COL_B1"}, matrix[0], "row 0 mismatch");
        assertArrayEquals(new String[]{"HIDDEN_ROW_A2", "HIDDEN_ROW_AND_COL_B2"}, matrix[1], "row 1 mismatch");
    }

    @Test
    void removeBlankRows_false_preservesMissingRowsAsFillerRows() throws Exception {
        byte[] workbookBytes = buildWorkbookBytes(wb -> {
            Sheet sheet = wb.createSheet("S1");

            // Row 0 present
            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("R0C0");

            // Row 1 intentionally NOT created (missing)

            // Row 2 present
            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("R2C0");
        });

        SheetConfig cfg = new SheetConfig(
                false, // removeBlankRows => should preserve missing row via fillMissingRows
                false,
                false,
                true,
                Set.of()
        );

        String[][] matrix = parseFirstSheet(workbookBytes, cfg);

        assertNotNull(matrix, "matrix must not be null");
        assertEquals(3, matrix.length, "missing row should be represented as a filler row when removeBlankRows=false");
        assertEquals("R0C0", matrix[0][0], "row 0 value mismatch");
        assertEquals("R2C0", matrix[2][0], "row 2 value mismatch");
    }

    @Test
    void removeBlankRows_true_doesNotAddFillerRowsForMissingRows() throws Exception {
        byte[] workbookBytes = buildWorkbookBytes(wb -> {
            Sheet sheet = wb.createSheet("S1");

            Row r0 = sheet.createRow(0);
            r0.createCell(0).setCellValue("R0C0");

            // Row 1 intentionally missing

            Row r2 = sheet.createRow(2);
            r2.createCell(0).setCellValue("R2C0");
        });

        SheetConfig cfg = new SheetConfig(
                true,  // removeBlankRows => fillMissingRows no-ops
                false,
                false,
                true,
                Set.of()
        );

        String[][] matrix = parseFirstSheet(workbookBytes, cfg);

        assertNotNull(matrix, "matrix must not be null");
        assertEquals(2, matrix.length, "missing row should not be synthesized when removeBlankRows=true");
        assertEquals("R0C0", matrix[0][0], "row 0 value mismatch");
        assertEquals("R2C0", matrix[1][0], "row 1 (originally row 2) value mismatch");
    }

    @Test
    void parse_withInvalidXml_throwsSaxException() throws Exception {
        SheetConfig cfg = new SheetConfig(
                false,
                false,
                false,
                true,
                Set.of()
        );

        // Build a real workbook just to obtain valid SharedStrings/StylesTable instances,
        // then feed invalid sheet XML to the parser.
        byte[] workbookBytes = buildWorkbookBytes(wb -> wb.createSheet("S1"));

        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(workbookBytes))) {
            XSSFReader reader = new XSSFReader(pkg);
            SharedStrings sharedStrings = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            SheetXMLReader parser = new SheetXMLReader(cfg, sharedStrings, styles);

            InputSource bad = new InputSource(new ByteArrayInputStream("<worksheet><row>".getBytes()));
            assertThrows(SAXException.class, () -> parser.parse(bad), "expected SAXException for malformed XML");
        }
    }

    // ---------------- helpers ----------------

    private static byte[] buildWorkbookBytes(Consumer<XSSFWorkbook> workbookCustomizer) throws IOException {
        try (XSSFWorkbook wb = new XSSFWorkbook()) {
            workbookCustomizer.accept(wb);
            try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
                wb.write(bos);
                return bos.toByteArray();
            }
        }
    }

    private static String[][] parseFirstSheet(byte[] workbookBytes, SheetConfig cfg) throws Exception {
        try (OPCPackage pkg = OPCPackage.open(new ByteArrayInputStream(workbookBytes))) {
            XSSFReader reader = new XSSFReader(pkg);

            SharedStrings sharedStrings = reader.getSharedStringsTable();
            StylesTable styles = reader.getStylesTable();

            try (InputStream sheetStream = firstSheetStream(reader)) {
                SheetXMLReader parser = new SheetXMLReader(cfg, sharedStrings, styles);
                parser.parse(new InputSource(sheetStream));
                return parser.getSheetData();
            }
        }
    }

    private static InputStream firstSheetStream(XSSFReader reader) throws IOException, InvalidFormatException {
        XSSFReader.SheetIterator it = (XSSFReader.SheetIterator) reader.getSheetIterator();
        assertTrue(it.hasNext(), "workbook must contain at least one sheet");
        return it.next();
    }
}