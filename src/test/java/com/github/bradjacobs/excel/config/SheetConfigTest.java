/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.config;

import com.github.bradjacobs.excel.sanitize.SanitizeType;
import com.github.bradjacobs.excel.sanitize.SpecialCharacterSanitizer;
import org.junit.jupiter.api.Test;

import java.util.EnumSet;
import java.util.Set;

import static com.github.bradjacobs.excel.sanitize.SanitizeType.BASIC_DIACRITICS;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.DASHES;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.QUOTES;
import static com.github.bradjacobs.excel.sanitize.SanitizeType.SPACES;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotSame;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

// dev note: this test class was 97% AI-generated and it did pretty well.
class SheetConfigTest {

    @Test
    void builderCreatesConfigWithDefaultValues() {
        SheetConfig config = SheetConfig.builder().build();

        assertFalse(config.skipBlankRows());
        assertFalse(config.skipBlankColumns());
        assertFalse(config.skipHiddenCells());
        assertTrue(config.trimStringValues());
        assertEquals(SpecialCharacterSanitizer.DEFAULT_FLAGS, config.getCharSanitizeFlags());
    }

    @Test
    void builderSetsAllBooleanConfigValuesToTrue() {
        SheetConfig config = SheetConfig.builder()
                .skipBlankRows(true)
                .skipBlankColumns(true)
                .skipHiddenCells(true)
                .trimStringValues(true)
                .build();

        assertTrue(config.skipBlankRows());
        assertTrue(config.skipBlankColumns());
        assertTrue(config.skipHiddenCells());
        assertTrue(config.trimStringValues());
    }

    @Test
    void builderSetsAllBooleanConfigValuesToFalse() {
        SheetConfig config = SheetConfig.builder()
                .skipBlankRows(false)
                .skipBlankColumns(false)
                .skipHiddenCells(false)
                .trimStringValues(false)
                .build();

        assertFalse(config.skipBlankRows());
        assertFalse(config.skipBlankColumns());
        assertFalse(config.skipHiddenCells());
        assertFalse(config.trimStringValues());
    }

    @Test
    void disableAllSanitationClearsSanitizeTypesAndDisablesTrimming() {
        SheetConfig config = SheetConfig.builder()
                .skipBlankRows(true)
                .skipBlankColumns(true)
                .skipHiddenCells(true)
                .disableAllSanitation()
                .build();

        assertTrue(config.skipBlankRows());
        assertTrue(config.skipBlankColumns());
        assertTrue(config.skipHiddenCells());
        assertFalse(config.trimStringValues());
        assertTrue(config.getCharSanitizeFlags().isEmpty());
    }

    @Test
    void sanitizeFlagsCanBeIndividuallyEnabled() {
        SheetConfig config = SheetConfig.builder()
                .disableAllSanitation()
                .sanitizeSpaces(true)
                .sanitizeQuotes(true)
                .sanitizeDiacritics(true)
                .sanitizeDashes(true)
                .build();

        assertEquals(
                EnumSet.of(SPACES, QUOTES, BASIC_DIACRITICS, DASHES),
                config.getCharSanitizeFlags());
    }

    @Test
    void sanitizeFlagsCanBeIndividuallyDisabled() {
        SheetConfig config = SheetConfig.builder()
                .sanitizeSpaces(false)
                .sanitizeQuotes(false)
                .sanitizeDiacritics(false)
                .sanitizeDashes(false)
                .build();

        assertTrue(config.getCharSanitizeFlags().isEmpty());
    }

    @Test
    void sanitizeFlagsCanBeToggledOffAndBackOn() {
        SheetConfig config = SheetConfig.builder()
                .sanitizeSpaces(false)
                .sanitizeQuotes(false)
                .sanitizeDiacritics(false)
                .sanitizeDashes(false)
                .sanitizeSpaces(true)
                .sanitizeQuotes(true)
                .sanitizeDiacritics(true)
                .sanitizeDashes(true)
                .build();

        assertEquals(
                EnumSet.of(SPACES, QUOTES, BASIC_DIACRITICS, DASHES),
                config.getCharSanitizeFlags());
    }

    @Test
    void trimCanBeReEnabledAfterDisablingAllSanitation() {
        SheetConfig config = SheetConfig.builder()
                .disableAllSanitation()
                .trimStringValues(true)
                .build();

        assertTrue(config.trimStringValues());
        assertTrue(config.getCharSanitizeFlags().isEmpty());
    }

    @Test
    void returnedSanitizeFlagsAreImmutable() {
        SheetConfig config = SheetConfig.builder().build();

        Set<SanitizeType> sanitizeFlags = config.getCharSanitizeFlags();

        assertThrows(UnsupportedOperationException.class, () -> sanitizeFlags.add(DASHES));
    }

    @Test
    void builtConfigUsesDefensiveCopyOfBuilderSanitizeFlags() {
        SheetConfig.Builder builder = SheetConfig.builder()
                .sanitizeSpaces(true)
                .sanitizeQuotes(true)
                .sanitizeDashes(false)
                .sanitizeDiacritics(false);

        SheetConfig firstConfig = builder.build();

        builder.disableAllSanitation()
                .sanitizeDashes(true)
                .sanitizeDiacritics(true);

        SheetConfig secondConfig = builder.build();

        assertEquals(EnumSet.of(SPACES, QUOTES), firstConfig.getCharSanitizeFlags());
        assertEquals(EnumSet.of(DASHES, BASIC_DIACRITICS), secondConfig.getCharSanitizeFlags());
    }

    @Test
    void builderMethodsReturnSameBuilderInstanceForFluentChaining() {
        SheetConfig.Builder builder = SheetConfig.builder();

        assertSame(builder, builder.skipBlankRows(true));
        assertSame(builder, builder.skipBlankColumns(true));
        assertSame(builder, builder.skipHiddenCells(true));
        assertSame(builder, builder.trimStringValues(false));
        assertSame(builder, builder.sanitizeSpaces(false));
        assertSame(builder, builder.sanitizeQuotes(false));
        assertSame(builder, builder.sanitizeDiacritics(true));
        assertSame(builder, builder.sanitizeDashes(true));
        assertSame(builder, builder.disableAllSanitation());
    }

    @Test
    void builderFactoryReturnsNewBuilderInstances() {
        SheetConfig.Builder firstBuilder = SheetConfig.builder();
        SheetConfig.Builder secondBuilder = SheetConfig.builder();

        assertNotSame(firstBuilder, secondBuilder);
    }

    @Test
    void sheetConfigSanitizeFlagsUnmodifiable() {
        SheetConfig config = SheetConfig.builder().build();
        Set<SanitizeType> flags = config.getCharSanitizeFlags();
        assertThrows(UnsupportedOperationException.class, () -> {
            flags.add(DASHES);
        });
    }
}
