package com.incesoft.tools.excel;

import java.awt.Color;
import java.io.File;

import com.incesoft.tools.excel.support.CellFormat;

public class Test {

	public static void main(String[] args) {

		WriterSupport wxs = WriterSupport.newInstance(new File("out.xlsx"));
		wxs.open();
		wxs.increaseRow();
		for (int i = 0; i < 5; i++) {
			wxs.increaseRow();
			wxs.writeRow(new String[] { "floydd" + i }, new CellFormat[] { new CellFormat(
					(i % 2 == 0 ? Color.PINK.getRGB() : Color.GREEN.getRGB()), -1, 0) });
		}
		wxs.close();
	}

}
