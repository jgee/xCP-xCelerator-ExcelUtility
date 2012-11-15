package com.emc.community.xcelerators.excelutil.impl;

import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Workbook;

import com.emc.community.xcelerators.excelutil.params.FormattingOptions;

/**
 * Abstracts formatting of the value that is supposed to be set in the cell
 * @author Michal.Malczewski
 *
 */
abstract class ValueSetter {

	/**
	 * Sets value in of a cell
	 * @param cell
	 * @param value
	 */
	abstract void setValue(Cell cell, String value) throws Exception;
	
	/**
	 * Gets the default ValueSetter if no formatting options are provided
	 * @return
	 */
	public static ValueSetter getDefault() {
		return new StringValueSetter();
	}
	
	/**
	 * Gets appropriate value setter for the formatting options provided
	 * @param options
	 * @return
	 */
	public static ValueSetter getValueSetter(FormattingOptions options, Workbook workbook) {
		switch (options.getCellType()) {
		case bool:
			return new BooleanValueSetter();
		case date:
			return new DateValueSetter(options.getSourceDateFormat(), options.getOutputDateFormat(), workbook);
		case number:
			return new DoubleValueSetter();
		default:
			throw new IllegalStateException("Unsupported cell type: " + options.getCellType());
		}
	}
	
	
	private static class StringValueSetter extends ValueSetter {

		public void setValue(Cell cell, String value) throws Exception {
			cell.setCellValue(value);
		}
		
	}
	
	/**
	 * Sets {@link Boolean} values
	 */
	private static class BooleanValueSetter extends ValueSetter {

		public void setValue(Cell cell, String value) throws Exception {
			boolean booleanValue = Boolean.valueOf(value);
			cell.setCellValue(booleanValue);
		}
	}

	/**
	 * Sets {@link Double} values
	 */
	private static class DoubleValueSetter extends ValueSetter {

		public void setValue(Cell cell, String value) throws Exception {
			double doubleValue = Double.parseDouble(value);
			cell.setCellValue(doubleValue);
		}
	}
	
	/**
	 * Sets {@link Date} values
	 */
	private static class DateValueSetter extends ValueSetter {
		
		private final SimpleDateFormat 	dateFormat;
		private final CellStyle		cellStyle;
		
		public DateValueSetter(String inputDateFormat, String outputDateFormat, Workbook workbook) {
			if (inputDateFormat == null) {
				throw new IllegalArgumentException("No dateFormatString provided for column");
			}
			dateFormat = new SimpleDateFormat(inputDateFormat);
			dateFormat.setLenient(false);
			
			if (outputDateFormat == null ){
				throw new IllegalArgumentException("No outputDateFormat provided for column");
			}
			
			cellStyle = workbook.createCellStyle();
			DataFormat formatter = workbook.createDataFormat();
			cellStyle.setDataFormat(formatter.getFormat(outputDateFormat));
		}

		public void setValue(Cell cell, String value) throws Exception {
			Date dateValue = dateFormat.parse(value);
			cell.setCellValue(dateValue);
			cell.setCellStyle(cellStyle);
		}
	}
}
