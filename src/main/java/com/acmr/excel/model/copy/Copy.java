package com.acmr.excel.model.copy;

import java.io.Serializable;

public class Copy implements Serializable{
	private String excelId;
	private String sheetId;
	private Orignal orignal;
	private Target target;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public String getSheetId() {
		return sheetId;
	}

	public void setSheetId(String sheetId) {
		this.sheetId = sheetId;
	}

	public Orignal getOrignal() {
		return orignal;
	}

	public void setOrignal(Orignal orignal) {
		this.orignal = orignal;
	}

	public Target getTarget() {
		return target;
	}

	public void setTarget(Target target) {
		this.target = target;
	}

}
