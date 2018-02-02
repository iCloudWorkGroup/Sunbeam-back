package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Border implements Serializable{
	private Boolean top = false;
	private Boolean right = false;
	private Boolean bottom = false;
	private Boolean left = false;
	private Boolean all = false;
	private Boolean outer = false;
	private Boolean none = false;

	public Boolean getTop() {
		return top;
	}

	public void setTop(Boolean top) {
		this.top = top;
	}

	public Boolean getRight() {
		return right;
	}

	public void setRight(Boolean right) {
		this.right = right;
	}

	public Boolean getBottom() {
		return bottom;
	}

	public void setBottom(Boolean bottom) {
		this.bottom = bottom;
	}

	public Boolean getLeft() {
		return left;
	}

	public void setLeft(Boolean left) {
		this.left = left;
	}

	public Boolean getAll() {
		return all;
	}

	public void setAll(Boolean all) {
		this.all = all;
	}

	public Boolean getOuter() {
		return outer;
	}

	public void setOuter(Boolean outer) {
		this.outer = outer;
	}

	public Boolean getNone() {
		return none;
	}

	public void setNone(Boolean none) {
		this.none = none;
	}

}
