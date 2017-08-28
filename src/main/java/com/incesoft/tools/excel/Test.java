package com.incesoft.tools.excel;

import java.awt.Color;
import java.io.File;

import com.incesoft.tools.excel.xlsx.CellFormat;
import com.incesoft.tools.excel.xlsx.XLSXWriterSupport;

public class Test {

	public static void main(String[] args) {

		XLSXWriterSupport wxs = new XLSXWriterSupport(new File("out.xlsx"));
		wxs.open();
		wxs.increaseRow();
		for (int i = 0; i < 5; i++) {
			wxs.increaseRow();
			wxs.writeRow(new String[] { "value" + i }, new CellFormat[] { new CellFormat(
					(i % 2 == 0 ? Color.PINK.getRGB() : Color.GREEN.getRGB()), -1, 0) });
		}
		wxs.close();
	}

}
