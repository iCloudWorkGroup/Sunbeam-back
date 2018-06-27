package com.acmr.excel.model;

import java.io.Serializable;

public class RowOrColExpand implements Serializable {
	/*增加类型，row或col*/
	private String type;
	/*增加的数量*/
	private int num;

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public int getNum() {
		return num;
	}

	public void setNum(int num) {
		this.num = num;
	}

}
