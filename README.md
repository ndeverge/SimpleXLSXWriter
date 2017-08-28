SimpleXLSXWriter - A very simple and efficient API for writing XLSX files
=======================================

> Credit: this project was first published in Google Code: https://code.google.com/p/sjxslx/ and it seems unmaintained.  It was then moved it to GitHub (https://github.com/davidpelfree/sjxlsx). I updated it in order to only write XLSX files, with no external dependencies.

It is a **simple and efficient** java tool for reading, writing and modifying XLSX. The most important purpose to code it is for performance consideration -- all the popular ones like POI sucks in both memory consuming and parse/write speed.

- memory

sjxlsx provides two modes (classic & stream) to read/modify sheets. In classic mode, all records of the sheet will be loaded. In stream mode (also named iterate mode), you can read record one after another which save a lot memory.

- speed

Microsoft XLSX use XML+zip (OOXML) to store the data. So, to be fast, sjxlsx use STAX for XML input and output. And I recommend the WSTX implementation of STAX (it's the fastest in my testing).

Sample code
-----------
```java
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
```
