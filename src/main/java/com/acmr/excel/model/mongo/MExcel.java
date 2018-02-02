package com.acmr.excel.model.mongo;

import java.io.Serializable;

public class MExcel implements Serializable {
	private String id;
	private int step;
	private String sheetName;
	private int maxrow;
	private int maxcol;

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

}
