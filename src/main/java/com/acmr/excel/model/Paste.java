package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class Paste implements Serializable {
	
	private List<OuterPasteData> pasteData;
	private int colLen;
	private int rowLen;


	public List<OuterPasteData> getPasteData() {
		return pasteData;
	}

	public void setPasteData(List<OuterPasteData> pasteData) {
		this.pasteData = pasteData;
	}

	public int getColLen() {
		return colLen;
	}

	public void setColLen(int colLen) {
		this.colLen = colLen;
	}

	public int getRowLen() {
		return rowLen;
	}

	public void setRowLen(int rowLen) {
		this.rowLen = rowLen;
	}

}
