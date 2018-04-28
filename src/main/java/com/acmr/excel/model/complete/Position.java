package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Position implements Serializable{
	private StrandX strandX = new StrandX();
	private StrandY strandY = new StrandY();
	private Max max = new Max();

	public StrandX getStrandX() {
		return strandX;
	}

	public void setStrandX(StrandX strandX) {
		this.strandX = strandX;
	}

	public StrandY getStrandY() {
		return strandY;
	}

	public void setStrandY(StrandY strandY) {
		this.strandY = strandY;
	}

	public Max getMax() {
		return max;
	}

	public void setMax(Max max) {
		this.max = max;
	}

}
