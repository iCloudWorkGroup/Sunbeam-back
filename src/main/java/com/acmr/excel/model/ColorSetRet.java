package com.acmr.excel.model;

import java.io.Serializable;

public class ColorSetRet implements Serializable {
	private int index;
	private String errorMessage;

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public String getErrorMessage() {
		return errorMessage;
	}

	public void setErrorMessage(String errorMessage) {
		this.errorMessage = errorMessage;
	}

}
