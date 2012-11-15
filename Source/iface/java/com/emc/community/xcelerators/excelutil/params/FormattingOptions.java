package com.emc.community.xcelerators.excelutil.params;


/**
 * Represents formatting options of values in cells.
 * Makes module's API self-descriptive in the Process Builder.
 * @author Michal.Malczewski
 *
 */
public class FormattingOptions {

	private boolean formatAsNumber;
	private boolean formatAsDate;
	private String sourceDateFormat;
	private String outputDateFormat;
	private boolean formatAsBoolean;
	
	private CellType cellType;
	public CellType getCellType() {
		return cellType;
	}
	
	public boolean isFormatAsNumber() {
		return formatAsNumber;
	}
	public void setFormatAsNumber(boolean formatAsNumber) {
		asserAllFalse();
		this.formatAsNumber = formatAsNumber;
		cellType = CellType.number;
	}
	
	public boolean isFormatAsDate() {
		return formatAsDate;
	}
	
	public void setFormatAsDate(boolean formatAsDate) {
		asserAllFalse();
		this.formatAsDate = formatAsDate;
		cellType = CellType.date;
	}
	
	public void setSourceDateFormat(String dateFormat) {
		this.sourceDateFormat = dateFormat;
	}
	
	public String getSourceDateFormat() {
		return sourceDateFormat;
	}
	
	public String getOutputDateFormat() {
		return outputDateFormat;
	}

	public void setOutputDateFormat(String outputDateFormat) {
		this.outputDateFormat = outputDateFormat;
	}

	public void setCellType(CellType cellType) {
		this.cellType = cellType;
	}

	public boolean isFormatAsBoolean() {
		return formatAsBoolean;
	}
	public void setFormatAsBoolean(boolean formatAsBoolean) {
		asserAllFalse();
		this.formatAsBoolean = formatAsBoolean;
		cellType = CellType.bool;
	}
	
	/**
	 * Only one value can be set to true
	 */
	private void asserAllFalse() {
		if (formatAsBoolean || formatAsNumber || formatAsDate) {
			throw new IllegalStateException("formatAs* already set");
		}
	}
	
	/**
	 * Represents format type
	 * @author Michal.Malczewski
	 *
	 */
	public static enum CellType  {
		string, number, date, bool
	}
}
