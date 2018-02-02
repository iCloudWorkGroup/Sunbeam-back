package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Max implements Serializable{
	private MaxStrandX maxStrandX;
	private MaxStrandY MaxStrandY;

	public MaxStrandX getMaxStrandX() {
		return maxStrandX;
	}

	public void setMaxStrandX(MaxStrandX maxStrandX) {
		this.maxStrandX = maxStrandX;
	}

	public MaxStrandY getMaxStrandY() {
		return MaxStrandY;
	}

	public void setMaxStrandY(MaxStrandY maxStrandY) {
		MaxStrandY = maxStrandY;
	}

}
