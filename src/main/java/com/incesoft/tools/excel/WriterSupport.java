package com.incesoft.tools.excel;

import com.incesoft.tools.excel.support.CellFormat;
import com.incesoft.tools.excel.support.XLSXWriterSupport;

import java.io.File;
import java.io.OutputStream;

abstract public class WriterSupport {


	protected File file;

	public void setFile(File file) {
		this.file = file;
	}

	protected OutputStream output;

	public void setOutputStream(OutputStream output) {
		this.output = output;
	}

	abstract protected int getMaxRowNumOfSheet();

	abstract public void open();

	abstract public void createNewSheet();

	abstract public void writeRow(String[] rowData);

	abstract public void writeRow(String[] rowData, CellFormat[] formats);

	abstract public void close();

	public static WriterSupport newInstance(File f) {
		WriterSupport support = null;
		support = new XLSXWriterSupport();
		support.setFile(f);
		return support;
	}

	public static WriterSupport newInstance(OutputStream outputStream) {
		WriterSupport support = null;
		support = new XLSXWriterSupport();
		support.setOutputStream(outputStream);
		return support;
	}

	public int getSheetIndex() {
		return sheetIndex;
	}

	public int getRowPos() {
		return rowpos;
	}

	protected int rowpos = getMaxRowNumOfSheet();

	protected int sheetIndex = -1;

	public void increaseRow() {
		rowpos++;

		if (rowpos > getMaxRowNumOfSheet()) {// 判断是否需要新建一个sheet
			sheetIndex++;
			createNewSheet();
			rowpos = -1;
		}
	}

}
