package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Border implements Serializable{
	private Integer all ;
	private Integer bottom ;
	private Integer left ;
	private Integer none ;
	private Integer outer ;
	private Integer right ;
	private Integer top ;
	public Integer getAll() {
		return all;
	}
	public void setAll(Integer all) {
		this.all = all;
	}
	public Integer getBottom() {
		return bottom;
	}
	public void setBottom(Integer bottom) {
		this.bottom = bottom;
	}
	public Integer getLeft() {
		return left;
	}
	public void setLeft(Integer left) {
		this.left = left;
	}
	public Integer getNone() {
		return none;
	}
	public void setNone(Integer none) {
		this.none = none;
	}
	public Integer getOuter() {
		return outer;
	}
	public void setOuter(Integer outer) {
		this.outer = outer;
	}
	public Integer getRight() {
		return right;
	}
	public void setRight(Integer right) {
		this.right = right;
	}
	public Integer getTop() {
		return top;
	}
	public void setTop(Integer top) {
		this.top = top;
	}

}
