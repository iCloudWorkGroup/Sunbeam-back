package com.acmr.excel.model.complete;

public class Rule {

	private String formula1;
	private String formula2;
	private Integer index;
	private Integer type;
	private int count;

	public String getFormula1() {
		return formula1;
	}

	public void setFormula1(String formula1) {
		this.formula1 = formula1;
	}

	public String getFormula2() {
		return formula2;
	}

	public void setFormula2(String formula2) {
		this.formula2 = formula2;
	}

	public Integer getIndex() {
		return index;
	}

	public void setIndex(Integer index) {
		this.index = index;
	}

	public Integer getType() {
		return type;
	}

	public void setType(Integer type) {
		this.type = type;
	}

	public int getCount() {
		return count;
	}

	public void setCount(int count) {
		this.count = count;
	}
}
