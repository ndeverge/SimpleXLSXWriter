package com.incesoft.tools.excel.xlsx;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.net.URISyntaxException;
import java.net.URL;

/**
 * @author floyd
 * 
 */
public class XLSXWriterSupport {

    private  File file;

    private OutputStream output;

	private SimpleXLSXWorkbook workbook;

    private int rowpos = getMaxRowNumOfSheet();

    private int sheetIndex = -1;

    public XLSXWriterSupport(OutputStream outputStream) {
        super();
        this.output = outputStream;
    }

    public XLSXWriterSupport(File file) {
        super();
        this.file = file;
    }

    public void open() {
        final URL emptyXLSX = getClass().getResource("/empty.xlsx");
        if (emptyXLSX == null) {
			throw new IllegalStateException("no empty.xlsx found in classpath");
		}
        try {
            workbook = new SimpleXLSXWorkbook(new File(emptyXLSX.toURI()));
        } catch (URISyntaxException e) {
            throw new IllegalStateException("no empty.xlsx found in classpath");
        }
    }

	private Sheet sheet;

	private int getMaxRowNumOfSheet() {
		return Integer.MAX_VALUE / 2;
	}

	public void writeRow(String[] rowData) {
		writeRow(rowData, null);
	}

	public void writeRow(String[] rowData, CellFormat[] formats) {
		for (int col = 0; col < rowData.length; col++) {
			String string = rowData[col];
			if (string == null)
				continue;
			CellFormat format = null;
			if (formats != null && formats.length > 0) {
				for (CellFormat cellFormat : formats) {
					if (cellFormat != null && cellFormat.getCellIndex() == col) {
						format = cellFormat;
						break;
					}
				}
			}
			CellStyle cellStyle = null;
			if (format != null && (format.getBackColor() != -1 || format.getForeColor() != -1)) {
				Font font = null;
				Fill fill = null;
				if (format.getForeColor() != -1) {
					font = workbook.createFont();
					font.setColor(format.getForeColor());
				}
				if (format.getBackColor() != -1) {
					fill = workbook.createFill();
					fill.setFgColor(format.getBackColor());
				}
				cellStyle = workbook.createStyle(font, fill);
			}
			sheet.modify(rowpos, col, string, cellStyle);
		}
	}

	public void close() {
		if (workbook == null)
			return;
		OutputStream fos = null;
		try {
			if (file != null) {
				fos = new FileOutputStream(file);
			} else if (output != null) {
				fos = output;
			} else {
				throw new IllegalStateException("no output specified");
			}
			workbook.commit(fos);
		} catch (Exception e) {
			throw new RuntimeException(e);
		} finally {
			if (fos != null)
				closeQuietly(fos);
			if (workbook != null)
				workbook.close();
		}
	}

	private void closeQuietly(OutputStream output) {
		try {
			if (output != null) {
				output.close();
			}
		} catch (IOException var2) {
        }

	}

	public void createNewSheet() {
		if (sheetIndex > 0) {
			throw new IllegalStateException("only one sheet allowed");
		}
		sheet = workbook.getSheet(sheetIndex, true);
	}

    public void increaseRow() {
        rowpos++;

        if (rowpos > getMaxRowNumOfSheet()) {// 判断是否需要新建一个sheet
            sheetIndex++;
            createNewSheet();
            rowpos = -1;
        }
    }
}
