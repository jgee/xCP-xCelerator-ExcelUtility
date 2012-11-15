package com.emc.community.xcelerators.excelutil.impl;

import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.util.CellReference;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

import com.documentum.fc.common.DfLogger;
import com.emc.community.xcelerators.excelutil.params.ReadCell;
import com.emc.community.xcelerators.excelutil.params.ReadColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.ReadingOptions;

/**
 * Provides means for reading Excel sheets
 * @author Michal.Malczewski
 *
 */
class ExcelReader extends AbstractExcelManipulator {

	private static final String DEFAULT_DATE_FORMAT = "yyyy-MM-dd";
	private String dateFormat = DEFAULT_DATE_FORMAT;
	private SimpleDateFormat dateFormatter;
	private int rowsToSkip;
	private boolean strictMode = true;
	
	/**
	 * Creates a dedicated instance for a workbook.
	 * @param inputSheet The workbook to read
	 * @param options Optional parameters for reading
	 * @throws Exception
	 */
	public ExcelReader(InputStream inputSheet, ReadingOptions options) throws Exception {
		super(inputSheet);
		if (options != null) {
			if (options.getDateFormat() != null ) {
				dateFormat = options.getDateFormat();
			}
			if (options.isStrictMode() != null) {
				strictMode = options.isStrictMode();
			}
			rowsToSkip = options.getRowsToSkip();
		}
	}
	
	/**
	 * Reads values from a particular sheet
	 * @param sheetName Name of the sheet. It must exist in the workbook, an exception is thrown otherwise
	 * @return
	 */
	public ReadColumnWithValues[] readSheet(String sheetName) {
		Sheet sheet = getSheetByName(sheetName);
		final int height = sheet.getLastRowNum();

		FormulaEvaluator evaluator = getWorkbook().getCreationHelper().createFormulaEvaluator();
		
		Map<Integer, List<ReadCell>> rowsByColumns = new HashMap<Integer, List<ReadCell>>();;

		for (int rowNum = rowsToSkip; rowNum <= height; rowNum++) {
			Row row = sheet.getRow(rowNum);
			if (row == null) {
				DfLogger.debug(this, "Row {0} was null", new Object[] { rowNum }, null);
				continue;
			}

			for (int cellNum = 0; cellNum < row.getLastCellNum(); cellNum++) {
				Cell cell = row.getCell(cellNum);
				String cellValue = getCellValue(cell, evaluator, strictMode);
				CellReference ref = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
				ReadCell rc = new ReadCell(ref.formatAsString(), cellValue);
				store(rowsByColumns, cellNum, rc, row.getLastCellNum());
				DfLogger.debug(this, "Read value " + rc.toString(), null,null);
			}
		}
		
		// now create result required by module's public interface
		int columnsRead = rowsByColumns.keySet().size();
		ReadColumnWithValues[] readColumns = new ReadColumnWithValues[columnsRead];
		for (int i = 0; i < readColumns.length; i++) {
			List<ReadCell> rowsInCol = rowsByColumns.get(i);
			if (rowsInCol == null) {
				continue;
			}
			readColumns[i] = new ReadColumnWithValues(rowsInCol.toArray(new ReadCell[rowsInCol.size()]));
		}
		return readColumns;
	}
	
	private void store(Map<Integer, List<ReadCell>> rowsByColumns, int cellNum, ReadCell rc, int maxHeight) {
		List<ReadCell> rowsInColumn = rowsByColumns.get(cellNum);
		if (rowsInColumn == null) {
			rowsInColumn = new ArrayList<ReadCell>(maxHeight);
			rowsByColumns.put(cellNum, rowsInColumn);
		}
		rowsInColumn.add(rc);
	}
	
	private String getCellValue(Cell cell, FormulaEvaluator evaluator, boolean strictMode) {
		switch (cell.getCellType()) {
			case Cell.CELL_TYPE_BLANK:
				return "";
			case Cell.CELL_TYPE_BOOLEAN:
				return "" + cell.getBooleanCellValue();
			case Cell.CELL_TYPE_ERROR:
				return "" + cell.getErrorCellValue();
			case Cell.CELL_TYPE_FORMULA:
				return evaluator.evaluate(cell).formatAsString();
			case Cell.CELL_TYPE_NUMERIC:
				// dates are stored as numbers
				if (HSSFDateUtil.isCellDateFormatted(cell)) {		
					return formatAsString(cell.getDateCellValue());
				}
				return "" + cell.getNumericCellValue();
			case Cell.CELL_TYPE_STRING:
				return cell.getStringCellValue();
			default:
				CellReference cr = new CellReference(cell.getRowIndex(), cell.getColumnIndex());
				if (strictMode) {
					throw new IllegalStateException("Unble to read " + cr.formatAsString() + " - cell type " + cell.getCellType() + " is not supported. Set strictMode to false to skip this error");
				}
				DfLogger.warn(this, "Unable to read value from " + cr.formatAsString(), null, null);
				return "";
		}
	}
	
	/**
	 * Formats date as string. Date format can be provided in {@link ExcelReader#ExcelReader(InputStream, ReadingOptions)} parameters
	 * @param date Date to format
	 */
	private String formatAsString(Date date) {
		if (dateFormatter == null) {
			dateFormatter = new SimpleDateFormat(dateFormat);
		}
		return dateFormatter.format(date);
	}
}
