/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.engine.objectmodel;

import com.github.bradjacobs.excel.reader.AbstractExcelReaderTest;

import java.nio.file.Path;

import static com.github.bradjacobs.excel.testutils.TestResourceUtil.getResourceFilePath;

class StandardExcelReaderTest extends AbstractExcelReaderTest<StandardExcelReader, StandardExcelReader.Builder> {

    @Override
    protected StandardExcelReader.Builder createBuilder() {
        return StandardExcelReader.builder();
    }

    private static final Path VALID_TEST_INPUT_PSWD_PATH = getResourceFilePath("test_data_w_pswd_1234.xlsx");
}
