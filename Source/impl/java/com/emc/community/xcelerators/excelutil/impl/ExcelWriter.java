package com.emc.community.xcelerators.excelutil.impl;

import java.io.InputStream;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.emc.community.xcelerators.excelutil.params.ColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.FormattingOptions;

/**
 * Performs write operations on an Excel sheet
 * @author Michal.Malczewski
 *
 */
class ExcelWriter extends AbstractExcelManipulator {
	
	
	public ExcelWriter(InputStream inputSheet) throws Exception {
		super(inputSheet);
	}


	private Map<String, ValueSetter> valueSettersForColumns = new HashMap<String, ValueSetter>();
	
	public void storeValuesInSheet(String sheetName, ColumnWithValues[] columns, int startInRow) throws Exception {
		Sheet workSheet = getSheetByName(sheetName); 
		// add columns 
		processColumns(workSheet, columns, startInRow);
	}
	
	/**
	 * Saves the sheet into the stream
	 */
	public void saveTo(OutputStream out, boolean recalculateFormulas) throws Exception {
		if (recalculateFormulas == true) {
			recalculateFormulas();
		}
		
		// store result in the output stream
		getWorkbook().write(out);
	}
	
	private void processColumns(Sheet workSheet, ColumnWithValues[] columns, int startInRow) throws Exception {
		if (columns == null) {
			/*figure out new logger
			DfLogger.debug(this, "There were no columns to process", null, null);
			*/
			return;
		}
		
		
		int maxHeight = getMaxHeight(columns);
		// assume that all columns will have the same number of values. This approach to processing seems to be most optimal
		for (int rownum = 0; rownum < maxHeight; rownum ++) {
			int targetRow = rownum + startInRow;
			Row excelRow = workSheet.getRow(targetRow);
			if (excelRow == null) {
				excelRow = workSheet.createRow(targetRow);
			}
			
			// process all columns; assume that different columns might have different size
			for (int colnum = 0; colnum < columns.length; colnum++) {
				ColumnWithValues col = columns[colnum];
				if (col == null) {
					/*figure out new logger
					DfLogger.trace(this, "Column {0} was null", new Object[] {colnum}, null);
					*/
					continue;
				}
				String[] valuesInColumn = col.getRows();
				if (valuesInColumn == null) {
					/*figure out new logger
					DfLogger.trace(this, "Column {0} had no rows", new Object[] {colnum}, null);
					*/
					continue;
				}
				
				if (valuesInColumn.length <= rownum) {
					/*figure out new logger
					 * DfLogger.trace(this, "Column {0} had less values than {1}: {2}", new Object[] {colnum, maxHeight, valuesInColumn.length}, null);
					 */
					continue;
				}
				String cellValue = valuesInColumn[rownum];
				if (cellValue == null) {
					/*figure out new logger
					 * DfLogger.trace(this, "Value for [{0};{1}] was null", new Object[] {colnum, rownum}, null);
					 */
					continue;
				}
				Cell excelCell = excelRow.getCell(colnum);
				if (excelCell == null ) {
					excelCell = excelRow.createCell(colnum);
				}
				/*figure out new logger
				 * DfLogger.trace(this, "Preparing to set value in [{0};{1}] to {2}", new Object[] {colnum, rownum, cellValue}, null);
				 */
				try {
				ValueSetter setter = getValueSetter(workSheet, colnum, col.getFormattingOptions());
				setter.setValue(excelCell, cellValue);
				}
				catch (Exception e) {
					throw new IllegalStateException("Error setting value in column " + colnum + " for row " + rownum, e);
				}
			}
		}
	}
	
	
	private ValueSetter getValueSetter(Sheet workSheet, int colnum, FormattingOptions formattingOptions) {
		final String key = workSheet.getSheetName() + "_" + colnum;
		ValueSetter setter = valueSettersForColumns.get(key);
		if (setter == null) {
			if (formattingOptions == null) {
				setter = ValueSetter.getDefault();
			}
			else {
				setter = ValueSetter.getValueSetter(formattingOptions, getWorkbook());
			}
			valueSettersForColumns.put(key, setter);
		}
		return setter;
	}


	/**
	 * Gets maximum height among all columns
	 * @param cols
	 * @return
	 */
	private int getMaxHeight(ColumnWithValues[] cols) {
		int max = -1;
		for (ColumnWithValues col : cols) {
			if (col == null) {
				continue;
			}
			Object[] rows = col.getRows();
			if (rows == null) {
				continue;
			}
			if (rows.length > max) {
				max = rows.length;
			}
		}
		return max;
	}

}
