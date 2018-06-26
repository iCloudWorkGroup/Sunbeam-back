package com.acmr.excel.model;

import java.io.Serializable;

public class Frozen implements Serializable {
	private String oprCol;
	private String oprRow;
	private String viewRow;
	private String viewCol;// 用于接收前台的传参

	private Integer row;
	private Integer col;

	public String getOprCol() {
		return oprCol;
	}

	public void setOprCol(String oprCol) {
		this.oprCol = oprCol;
	}

	public String getOprRow() {
		return oprRow;
	}

	public void setOprRow(String oprRow) {
		this.oprRow = oprRow;
	}

	public String getViewRow() {
		return viewRow;
	}

	public void setViewRow(String viewRow) {
		this.viewRow = viewRow;
	}

	public String getViewCol() {
		return viewCol;
	}

	public void setViewCol(String viewCol) {
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
