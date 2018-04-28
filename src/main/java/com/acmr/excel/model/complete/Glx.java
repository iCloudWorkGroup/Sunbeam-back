package com.acmr.excel.model.complete;

import java.io.Serializable;

public class Glx implements Serializable {
	private String aliasX;
	private int left;
	private int width;
	private int originWidth;
	private int index;
	private boolean hidden;
	private OperProp operProp = new OperProp();

	public OperProp getOperProp() {
		return operProp;
	}

	public void setOperProp(OperProp operProp) {
		this.operProp = operProp;
	}

	public int getOriginWidth() {
		return originWidth;
	}

	public void setOriginWidth(int originWidth) {
		this.originWidth = originWidth;
	}

	public String getAliasX() {
		return aliasX;
	}

	public void setAliasX(String aliasX) {
		this.aliasX = aliasX;
	}

	public int getLeft() {
		return left;
	}

	public void setLeft(int left) {
		this.left = left;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public int getIndex() {
		return index;
	}

	public void setIndex(int index) {
		this.index = index;
	}

	public boolean isHidden() {
		return hidden;
	}

	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

}
