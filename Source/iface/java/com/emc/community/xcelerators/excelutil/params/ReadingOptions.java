package com.emc.community.xcelerators.excelutil.params;

/**
 * Represents options for Excel sheet reading.
 * Makes module's API self-descriptive in the Process Builder.
 * @author Michal.Malczewski
 *
 */
public class ReadingOptions {

	private Boolean strictMode;
	private int rowsToSkip;
	private String dateFormat;
	
	/**
	 * Default .ctor
	 */
	public ReadingOptions() {}
	
	/**
	 * Inline .ctor, used mostly for tests
	 */
	public ReadingOptions(String dateFormat, int rowsToSkip, Boolean strictMode) {
		this.dateFormat = dateFormat;
		this.rowsToSkip = rowsToSkip;
		this.strictMode = strictMode;
	}
	
	public Boolean isStrictMode() {
		return strictMode;
	}
	public void setStrictMode(boolean strictMode) {
		this.strictMode = strictMode;
	}
	public int getRowsToSkip() {
		return rowsToSkip;
	}
	public void setRowsToSkip(int rowsToSkip) {
		this.rowsToSkip = rowsToSkip;
	}
	public String getDateFormat() {
		return dateFormat;
	}
	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}
}
