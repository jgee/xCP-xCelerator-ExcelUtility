package com.emc.community.xcelerators.excelutil.impl;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.Calendar;

import com.emc.community.xcelerators.excelutil.params.ColumnWithValues;
import com.emc.community.xcelerators.excelutil.params.FormattingOptions;

public abstract class ExcelWriterTest extends ExcelTestBase {

	private static final boolean STORE_RESULT_AS_FILE = true;

	public void testWritingValuesInEmptySheet() throws Exception {
		final String dateFormatString = "yyyy-MM-dd";
		final SimpleDateFormat dateFormat = new SimpleDateFormat(
				dateFormatString);
		ColumnWithValues[] columns = new ColumnWithValues[10];
		for (int colnum = 0; colnum < columns.length; colnum++) {
			String[] rows = new String[2 + colnum];
			FormattingOptions formattingOptions = null;

			switch (colnum) {
			case 1:
				// number
				formattingOptions = new FormattingOptions();
				formattingOptions.setFormatAsNumber(true);
				break;
			case 2:
				// date
				formattingOptions = new FormattingOptions();
				formattingOptions.setFormatAsDate(true);
				formattingOptions.setSourceDateFormat(dateFormatString);
				formattingOptions.setOutputDateFormat(dateFormatString);
				break;
			case 3:
				// boolean
				formattingOptions = new FormattingOptions();
				formattingOptions.setFormatAsBoolean(true);
				break;
			default:
				break;
			}

			ColumnWithValues col = new ColumnWithValues(rows, formattingOptions);
			columns[colnum] = col;

			col.setRows(rows);
			for (int j = 0; j < rows.length; j++) {
				String value;
				switch (colnum) {
				case 1:
					// number
					value = "" + colnum + j;
					break;
				case 2:
					// date
					Calendar now = Calendar.getInstance();
					now.add(Calendar.DAY_OF_MONTH, j);
					value = dateFormat.format(now.getTime());
					break;
				case 3:
					// boolean
					value = "" + (j % 2 == 0);
					break;
				default:
					value = colnum + "_" + j;
					;
				}
				rows[j] = value;
			}
		}

		InputStream testSheet = getTestWorkbook();
		OutputStream out = getTestOutput();
		ExcelWriter xw = new ExcelWriter(testSheet);
		xw.storeValuesInSheet("SHEET1", columns, 2);
		xw.saveTo(out, true);
		out.close();
	}

	private OutputStream getTestOutput() throws FileNotFoundException, IOException {
		if (STORE_RESULT_AS_FILE) {
			File out = File.createTempFile(getClass().getName(), "." + getTestSheetExtension().name());
			System.out.println(out.getAbsolutePath());
			return new FileOutputStream(out);
		} else {
			return new ByteArrayOutputStream();
		}
	}
}
