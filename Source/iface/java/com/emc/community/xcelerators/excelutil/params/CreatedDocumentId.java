package com.emc.community.xcelerators.excelutil.params;

/**
 * Represents id of the created object.
 * Makes module's API self-descriptive in the Process Builder.
 * @author Michal.Malczewski
 *
 */
public class CreatedDocumentId {

	private String createdDocumentId;

	/**
	 * Default .ctor
	 */
	public CreatedDocumentId() {}
	
	/**
	 * Creates instance 
	 * @param createdDocumentId r_object_id of the created object
	 */
	public CreatedDocumentId(String createdDocumentId) {
		this.createdDocumentId = createdDocumentId;
	}
	
	public String getCreatedDocumentId() {
		return createdDocumentId;
	}
	
	public void setCreatedDocumentId(String createdDocumentId) {
		this.createdDocumentId = createdDocumentId;
	}
	
}
