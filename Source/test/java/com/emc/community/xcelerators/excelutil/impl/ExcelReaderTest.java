package com.emc.community.xcelerators.excelutil.impl;

import com.emc.community.xcelerators.excelutil.params.ReadColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.ReadingOptions;


public abstract class ExcelReaderTest extends ExcelTestBase {
	
	private static final String TEST_SHEET_NAME = "READING_TEST";

	public void testReading_noskip() throws Exception {
		ExcelReader reader = new ExcelReader(getTestWorkbook(), null);
		ReadColumnWithValues[] result = reader.readSheet(TEST_SHEET_NAME);
		assertNotNull(result);
		assertEquals("DATES", result[0].getRows()[0].getValue());
		assertEquals("NUMBERS", result[1].getRows()[0].getValue());
		assertEquals("FORMULAS", result[2].getRows()[0].getValue());
		assertResults(result, 0);
	}
	
	public void testReading_skip1() throws Exception {
		ExcelReader reader = new ExcelReader(getTestWorkbook(), new ReadingOptions(null, 1, true));
		ReadColumnWithValues[] result = reader.readSheet(TEST_SHEET_NAME);
		assertNotNull(result);
		assertResults(result, 1);
	}
	
	private void assertResults(ReadColumnWithValues[] result, int rowsToSkip) {
		assertEquals("2010-01-02", result[0].getRows()[1 - rowsToSkip].getValue());
		assertEquals("1.0", result[1].getRows()[1 - rowsToSkip].getValue());
		assertEquals("6.0", result[2].getRows()[1 - rowsToSkip].getValue());
	}
}
