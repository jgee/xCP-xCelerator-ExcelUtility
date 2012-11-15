package com.emc.community.xcelerators.excelutil.params;

/**
 * Represents requirements for the new object.
 * Makes module's API self-descriptive in the Process Builder.
 * @author Michal.Malczewski
 *
 */
public class NewObjectParams {

	private String newObjectType;
	private String newObjectName;
	private String newObjectPrimaryFolder;
	public String getNewObjectType() {
		return newObjectType;
	}
	public void setNewObjectType(String newObjectType) {
		this.newObjectType = newObjectType;
	}
	public String getNewObjectName() {
		return newObjectName;
	}
	public void setNewObjectName(String newObjectName) {
		this.newObjectName = newObjectName;
	}
	public String getNewObjectPrimaryFolder() {
		return newObjectPrimaryFolder;
	}
	public void setNewObjectPrimaryFolder(String newObjectPrimaryFolder) {
		this.newObjectPrimaryFolder = newObjectPrimaryFolder;
	}
	
	
}
