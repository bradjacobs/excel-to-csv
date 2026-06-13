/*
 * This file is subject to the terms and conditions defined in the 'LICENSE' file.
 */
package com.github.bradjacobs.excel.architecture;

import com.tngtech.archunit.core.domain.JavaClass;
import com.tngtech.archunit.core.domain.JavaClasses;
import com.tngtech.archunit.core.importer.ClassFileImporter;
import com.tngtech.archunit.core.importer.ImportOption;
import com.tngtech.archunit.library.dependencies.SliceAssignment;
import com.tngtech.archunit.library.dependencies.SliceIdentifier;
import com.tngtech.archunit.library.dependencies.SlicesRuleDefinition;
import org.junit.jupiter.api.Test;

class NoPackageCyclesTest {

    /**
     * Checks for any package cycles in the main project code.
     * note: this may turn out to be 'too aggressive', and
     * and be tuned later if/when needed.
     */
    @Test
    void everyPackageMustBeCycleFree() {
        JavaClasses classes = new ClassFileImporter()
                .withImportOption(ImportOption.Predefined.DO_NOT_INCLUDE_TESTS)
                .importPackages("com.github.bradjacobs.excel");

        SliceAssignment byExactPackage = new SliceAssignment() {
            @Override
            public SliceIdentifier getIdentifierOf(JavaClass javaClass) {
                return SliceIdentifier.of(javaClass.getPackageName());
            }

            @Override
            public String getDescription() {
                return "one slice per Java package";
            }
        };

        SlicesRuleDefinition.slices()
                .assignedFrom(byExactPackage)
                .should()
                .beFreeOfCycles()
                .check(classes);
    }
}
