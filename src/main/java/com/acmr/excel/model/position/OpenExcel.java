package com.acmr.excel.model.position;

import java.util.ArrayList;
import java.util.List;

public class OpenExcel {
	private String excelId;
	private int top;
	private int bottom;
	private int left;
	private int right;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public int getTop() {
		return top;
	}

	public void setTop(int top) {
		this.top = top;
	}

	public int getBottom() {
		return bottom;
	}

	public void setBottom(int bottom) {
		this.bottom = bottom;
	}

	public int getLeft() {
		return left;
	}

	public void setLeft(int left) {
		this.left = left;
	}

	public int getRight() {
		return right;
	}

	public void setRight(int right) {
		this.right = right;
	}
	
	public static void main(String[] args) {
		List<String> arr = new ArrayList<>(100);
		arr.add(60,"a");
	}
	
	
	

}
