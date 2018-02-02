package com.acmr.excel.model;

import java.io.Serializable;

public class ColWidth implements Serializable {
	private int col;
	private int offset;


	public int getOffset() {
		return offset;
	}

	public void setOffset(int offset) {
		this.offset = offset;
	}

	public int getCol() {
		return col;
	}

	public void setCol(int col) {
		this.col = col;
	}

	
}
