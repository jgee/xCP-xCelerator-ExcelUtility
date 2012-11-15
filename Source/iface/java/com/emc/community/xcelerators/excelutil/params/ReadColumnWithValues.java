package com.emc.community.xcelerators.excelutil.params;

import java.util.Arrays;

/**
 * Represents column read from a sheet
 * @author Michal.Malczewski
 *
 */
public class ReadColumnWithValues {

	private ReadCell[] rows;

	public ReadColumnWithValues() {}
	
	public ReadColumnWithValues(ReadCell[] rows) {
		this.rows = rows;
	}

	public ReadCell[] getRows() {
		return rows;
	}

	public void setRows(ReadCell[] rows) {
		this.rows = rows;
	}
	
	
	@Override
	public String toString() {
		return Arrays.toString(rows);
	}
	
}
