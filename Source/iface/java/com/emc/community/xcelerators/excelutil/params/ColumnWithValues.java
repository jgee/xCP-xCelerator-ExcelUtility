package com.emc.community.xcelerators.excelutil.params;

import java.util.Arrays;

/**
 * Represents values in an Excel column.
 * Makes module's API self-descriptive in the Process Builder. 
 * @author Michal.Malczewski
 *
 */
public class ColumnWithValues {
	
	private String[] rows;
	private FormattingOptions formattingOptions;
	
	/**
	 * Default .ctor
	 */
	public ColumnWithValues() {}

	/**
	 * Inline .ctor, used primarly for tests
	 * @param columnsInRow
	 */
	public ColumnWithValues(String[] columnsInRow, FormattingOptions formattingOptions) {
		this.rows = columnsInRow;
		this.formattingOptions = formattingOptions;
	}

	public String[] getRows() {
		return rows;
	}

	public void setRows(String[] columnsInRow) {
		this.rows = columnsInRow;
	}
	
	public FormattingOptions getFormattingOptions() {
		return formattingOptions;
	}
	
	public void setFormattingOptions(FormattingOptions formattingOptions) {
		this.formattingOptions = formattingOptions;
	}
	
	
	@Override
	public String toString() {
		// hashCode and equals not overriden, no need for that
		return Arrays.toString(rows);
	}
}
