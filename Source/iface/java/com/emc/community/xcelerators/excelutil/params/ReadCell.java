package com.emc.community.xcelerators.excelutil.params;

/**
 * Represents cell read from a sheet
 * @author Michal.Malczewski
 *
 */
public class ReadCell {

	private final String address;
	private final String value;
	
	public ReadCell(String address, String value) {
		this.address = address;
		this.value = value;
	}

	public String getAddress() {
		return address;
	}

	public String getValue() {
		return value;
	}
	
	@Override
	public String toString() {
		return address + ": " + value;
	}
}
