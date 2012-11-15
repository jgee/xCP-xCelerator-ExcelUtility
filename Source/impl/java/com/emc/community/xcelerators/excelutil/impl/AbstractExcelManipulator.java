package com.emc.community.xcelerators.excelutil.impl;

import java.io.InputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import com.documentum.fc.common.DfLogger;

/**
 * Base class for Excel manipulation classes
 * @author Michal.Malczewski
 *
 */
abstract class AbstractExcelManipulator {

	private final Workbook workbook;
	
	/**
	 * Default .ctor
	 * @param inputSheet Input stream to an Excel sheet
	 */
	public AbstractExcelManipulator(InputStream inputSheet) throws Exception {
		this.workbook = WorkbookFactory.create(inputSheet);
		//new XSSFWorkbook(inputSheet);
	}
	
	protected Workbook getWorkbook() {
		return workbook;
	}
	
	/**
	 * Gets sheet by name. Throws an exception if it does not exist.
	 * @param sheetName Name of the sheet to retrieve
	 */
	protected Sheet getSheetByName(String sheetName) {
		Sheet sheet = getWorkbook().getSheet(sheetName);
		if (sheet == null) {
			throw new IllegalStateException("Sheet " + sheetName + " not found");
		}
		return sheet;
	}
	
	/**
	 * Recalculates all formulas in the workbook
	 */
	protected void recalculateFormulas() {
		DfLogger.trace(this, "Recalculating formulas", null, null);
		FormulaEvaluator evaluator = getWorkbook().getCreationHelper().createFormulaEvaluator();
		for(int sheetNum = 0; sheetNum < getWorkbook().getNumberOfSheets(); sheetNum++) {
		    Sheet sheet = getWorkbook().getSheetAt(sheetNum);
		    for(Row r : sheet) {
		        for(Cell c : r) {
		            if(c.getCellType() == Cell.CELL_TYPE_FORMULA) {
		                evaluator.evaluateFormulaCell(c);
		            }
		        }
		    }
		}	
	}

}
