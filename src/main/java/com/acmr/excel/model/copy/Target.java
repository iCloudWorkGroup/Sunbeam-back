package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Target implements Serializable {
	private int oprCol;
	private int oprRow;

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

}
