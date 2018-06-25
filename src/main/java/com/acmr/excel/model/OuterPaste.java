package com.acmr.excel.model;

import java.io.Serializable;

public class OuterPaste implements Serializable {
	private int oprCol;
	private int oprRow;
	private String content;

	public int getOprCol() {
		return oprCol;
	}

	public void setOprCol(int oprCol) {
		this.oprCol = oprCol;
	}

	public int getOprRow() {
		return oprRow;
	}

	public void setOprRow(int oprRow) {
		this.oprRow = oprRow;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

}
