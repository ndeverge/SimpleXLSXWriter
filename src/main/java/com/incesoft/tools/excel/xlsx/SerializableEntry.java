package com.incesoft.tools.excel.xlsx;


import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamWriter;

/**
 * @author floyd
 * 
 */
public interface SerializableEntry {
	void serialize(XMLStreamWriter writer) throws XMLStreamException;
}
