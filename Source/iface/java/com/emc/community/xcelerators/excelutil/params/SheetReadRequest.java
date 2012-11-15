package com.emc.community.xcelerators.excelutil.params;

/**
 *  Represents required parameters for the Excel sheet reading operation.
 *  Makes module's API self-descriptive in the Process Builder. 
 * @author Michal.Malczewski
 *
 */
public class SheetReadRequest {

	private String workbookObjectId;
	private String sheetName;
	public String getWorkbookObjectId() {
		return workbookObjectId;
	}
	public void setWorkbookObjectId(String workbookObjectId) {
		this.workbookObjectId = workbookObjectId;
	}
	public String getSheetName() {
		return sheetName;
	}
	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}
	
	
}
