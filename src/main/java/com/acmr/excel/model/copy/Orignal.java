package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Orignal implements Serializable {
	private int endCol;
	private int endRow;
	private int startCol;
	private int startRow;
	public int getEndCol() {
		return endCol;
	}
	public void setEndCol(int endCol) {
		this.endCol = endCol;
	}
	public int getEndRow() {
		return endRow;
	}
	public void setEndRow(int endRow) {
		this.endRow = endRow;
	}
	public int getStartCol() {
		return startCol;
	}
	public void setStartCol(int startCol) {
		this.startCol = startCol;
	}
	public int getStartRow() {
		return startRow;
	}
	public void setStartRow(int startRow) {
		this.startRow = startRow;
	}

	
}
