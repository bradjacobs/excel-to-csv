/*
 * This file is subject to the terms and conditions defined in 'LICENSE' file.
 */
package com.github.bradjacobs.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.junit.jupiter.api.Test;

import java.util.HashSet;
import java.util.Set;

import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.SpecialCharacterSanitizer.CharSanitizeFlag.QUOTES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.mockito.Mockito.mock;
import static org.mockito.Mockito.when;

/**
 * Basic UnitTests for the CellValueReader.
 */
// NOTE: limited tests here because don't want to learn
//   about mocking all the different internal variations of cell types.
public class CellValueReaderTest {

    @Test
    public void withTrim() {
        CellValueReader cellValueReader = new CellValueReader(true, new HashSet<>());
        String inputString = "  dog  ";
        String expectedString = inputString.trim();
        Cell cell = createMockStringCell(inputString);
        String result = cellValueReader.getCellValue(cell);
        assertEquals(expectedString, result, "mismatch expected cell value string");
    }

    @Test
    public void withoutTrim() {
        CellValueReader cellValueReader = new CellValueReader(false, new HashSet<>());
        String inputString = "  dog  ";
        Cell cell = createMockStringCell(inputString);
        String result = cellValueReader.getCellValue(cell);
        assertEquals(inputString, result, "mismatch expected cell value string");
    }

    @Test
    public void withDiacriticsSet() {
        Set<SpecialCharacterSanitizer.CharSanitizeFlag> flagSet = Set.of(BASIC_DIACRITICS);
        CellValueReader cellValueReader = new CellValueReader(true, flagSet);
        Cell cell = createMockStringCell("Façade");
        String result = cellValueReader.getCellValue(cell);
        assertEquals("Facade", result, "mismatch expected cell value");
    }

    @Test
    public void nullCellValue() {
        CellValueReader cellValueReader = new CellValueReader(true, Set.of(QUOTES));
        String result = cellValueReader.getCellValue(null);
        assertEquals("", result, "mismatch expected cell value");
    }

    @Test
    public void formulaConversion() {
        // note: this is a semi-poor representation of a true Excel Formula cell.
        CellValueReader cellValueReader = new CellValueReader(true, Set.of(QUOTES));
        Cell cell = mock(Cell.class);
        when(cell.getCellType()).thenReturn(CellType.FORMULA);
        when(cell.getCachedFormulaResultType()).thenReturn(CellType.NUMERIC);
        when(cell.getCellFormula()).thenReturn("A1+B1"); // cellFormula string is only used if the code is misconfigured
        when(cell.getNumericCellValue()).thenReturn(31.2d);

        String result = cellValueReader.getCellValue(cell);
        assertEquals("31.2", result, "mismatch expected cell value");
    }

    @Test
    public void nullSanitizeFlags() {
        // sanitize flags parameter must be non-null
        Exception exception = assertThrows(IllegalArgumentException.class, () -> {
            new CellValueReader(true, null);
        });
        assertEquals("Must provide non-null charSanitizeFlags.", exception.getMessage(), "Mismatch exception message");
    }

    /**
     * Helper method for creating a Mock Cell input
     * @param cellValue string value for cell
     * @return cell
     */
    private Cell createMockStringCell(String cellValue) {
        Cell cell = mock(Cell.class);
        when(cell.getCellType()).thenReturn(CellType.STRING);
        when(cell.getRichStringCellValue()).thenReturn(new XSSFRichTextString(cellValue));
        return cell;
    }
}
