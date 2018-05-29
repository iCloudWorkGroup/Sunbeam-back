package com.acmr.excel.model;

import java.io.Serializable;

public class Frozen implements Serializable {
	private int oprCol;
	private int oprRow;
	private int viewRow;
	private int viewCol;//用于接收前台的传参
	
	private Integer row;
	private Integer col;

	
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

	public int getViewRow() {
		return viewRow;
	}

	public void setViewRow(int viewRow) {
		this.viewRow = viewRow;
	}

	public int getViewCol() {
		return viewCol;
	}

	public void setViewCol(int viewCol) {
		this.viewCol = viewCol;
	}

	public Integer getRow() {
		return row;
	}

	public void setRow(Integer row) {
		this.row = row;
	}

	public Integer getCol() {
		return col;
	}

	public void setCol(Integer col) {
		this.col = col;
	}


}
