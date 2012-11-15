package com.emc.community.xcelerators.excelutil.params;

import java.util.Arrays;

/**
 * Represents requirements for Excel manipulation. 
 * Makes module's API self-descriptive in the Process Builder.
 * @author Michal.Malczewski
 *
 */
public class StoreValuesParams {

	private String sourceDocumentId;
	private String sheetName;
	private ColumnWithValues[] columns;
	private int startWritingInRow;
	private boolean recalculateFormulas;
	
	/**
	 * Defialt .ctor
	 */
	public StoreValuesParams() {}
	
	public StoreValuesParams(String sourceDocumentId, String sheetName,	ColumnWithValues[] rows, int startWritingInRow, boolean recalculateFormulas) {
		this.sourceDocumentId 		= sourceDocumentId;
		this.sheetName 				= sheetName;
		this.columns 				= rows;
		this.startWritingInRow		= startWritingInRow;
		this.recalculateFormulas	= recalculateFormulas;
	}

	public String getSourceDocumentId() {
		return sourceDocumentId;
	}

	public void setSourceDocumentId(String sourceDocumentId) {
		this.sourceDocumentId = sourceDocumentId;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public ColumnWithValues[] getColumns() {
		return columns;
	}

	public void setColumns(ColumnWithValues[] rows) {
		this.columns = rows;
	}
	
	public int getStartWritingInRow() {
		return startWritingInRow;
	}

	public void setStartWritingInRow(int startInRow) {
		this.startWritingInRow = startInRow;
	}
	
	
	
	public boolean isRecalculateFormulas() {
		return recalculateFormulas;
	}

	public void setRecalculateFormulas(boolean recalculateFormulas) {
		this.recalculateFormulas = recalculateFormulas;
	}

	@Override
	public String toString() {
		return "docid: " + sourceDocumentId + "; sheet name: " + sheetName + "; columns: " + Arrays.toString(columns);
	}
	
	@Override
	public int hashCode() {
		return toString().hashCode();
	}
}
