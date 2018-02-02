package com.acmr.excel.model.history;

import java.io.Serializable;

public class ChangeArea implements Serializable{
	private int colIndex;
	private int rowIndex;
	private Object originalValue;
	private Object updateValue;
	private boolean isExist = true;

	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public Object getOriginalValue() {
		return originalValue;
	}

	public void setOriginalValue(Object originalValue) {
		this.originalValue = originalValue;
	}

	public Object getUpdateValue() {
		return updateValue;
	}

	public void setUpdateValue(Object updateValue) {
		this.updateValue = updateValue;
	}

	public boolean isExist() {
		return isExist;
	}

	public void setExist(boolean isExist) {
		this.isExist = isExist;
	}
	
}
