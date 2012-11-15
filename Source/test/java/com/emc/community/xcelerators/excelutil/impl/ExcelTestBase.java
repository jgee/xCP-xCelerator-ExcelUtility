package com.emc.community.xcelerators.excelutil.impl;

import java.io.InputStream;

import junit.framework.TestCase;

public abstract class ExcelTestBase extends TestCase {

	private InputStream getTestResource(String name) {
		InputStream is = getClass().getResourceAsStream("/" + name);
		if (is == null) {
			throw new IllegalStateException(name + " not found");
		}
		return is;
	}
	
	
	protected InputStream getTestWorkbook() {
		return getTestResource("test." + getTestSheetExtension().name());
	}
	
	protected abstract TestSheetExtension getTestSheetExtension();
	
	
	protected enum TestSheetExtension {
		xls, xlsx;
	}
	
}
