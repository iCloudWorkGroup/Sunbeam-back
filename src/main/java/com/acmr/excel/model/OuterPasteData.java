package com.acmr.excel.model;

import java.io.Serializable;

public class OuterPasteData implements Serializable {
	private int col;
	private int row;
	private String content;

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	public int getRow() {
		return row;
	}

	public void setRow(int row) {
		this.row = row;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

}
