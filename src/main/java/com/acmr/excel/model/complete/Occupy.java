package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Occupy implements Serializable {
	private List<String> col = new ArrayList<String>();
	private List<String> row = new ArrayList<String>();
	
	public List<String> getCol() {
		return col;
	}
	public void setCol(List<String> col) {
		this.col = col;
	}
	public List<String> getRow() {
		return row;
	}
	public void setRow(List<String> row) {
		this.row = row;
	}
	
}
