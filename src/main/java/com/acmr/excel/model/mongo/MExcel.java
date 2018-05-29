package com.acmr.excel.model.mongo;

import java.io.Serializable;

public class MExcel implements Serializable {
	private String id;
	private int step;
	private String sheetName;
	private int maxrow;
	private int maxcol;
	private String viewRowAlias;//可视行
	private String viewColAlias;//可视列
	private String colAlias;//冻结列
	private String rowAlias;//冻结行
	private boolean freeze;//是否冻结
		

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public int getStep() {
		return step;
	}

	public void setStep(int step) {
		this.step = step;
	}

	public String getSheetName() {
		return sheetName;
	}

	public void setSheetName(String sheetName) {
		this.sheetName = sheetName;
	}

	public int getMaxrow() {
		return maxrow;
	}

	public void setMaxrow(int maxrow) {
		this.maxrow = maxrow;
	}

	public int getMaxcol() {
		return maxcol;
	}

	public void setMaxcol(int maxcol) {
		this.maxcol = maxcol;
	}

	public String getViewRowAlias() {
		return viewRowAlias;
	}

	public void setViewRowAlias(String viewRowAlias) {
		this.viewRowAlias = viewRowAlias;
	}

	public String getViewColAlias() {
		return viewColAlias;
	}

	public void setViewColAlias(String viewColAlias) {
		this.viewColAlias = viewColAlias;
	}

	public String getColAlias() {
		return colAlias;
	}

	public void setColAlias(String colAlias) {
		this.colAlias = colAlias;
	}

	public String getRowAlias() {
		return rowAlias;
	}

	public void setRowAlias(String rowAlias) {
		this.rowAlias = rowAlias;
	}

	public boolean isFreeze() {
		return freeze;
	}

	public void setFreeze(boolean freeze) {
		this.freeze = freeze;
	}

}
