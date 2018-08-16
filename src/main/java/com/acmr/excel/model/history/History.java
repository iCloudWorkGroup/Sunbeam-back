package com.acmr.excel.model.history;

import java.io.Serializable;

public class History implements Serializable{
	
	// 记录方法名称
	private String name;
	//数据集id(表名)
	private String excelId;

	private Before  before = new Before();

	private Record record = new Record();

	public Before getBefore() {
		return before;
	}

	public void setBefore(Before before) {
		this.before = before;
	}

	public Record getRecord() {
		return record;
	}

	public void setRecord(Record record) {
		this.record = record;
	}

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

}
