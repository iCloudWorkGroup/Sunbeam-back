package com.acmr.mq;

import java.io.Serializable;

import javax.servlet.ServletInputStream;

public class Model implements Serializable {
	private int step;
	private int reqPath;
	private Object object;
	private String excelId;

	
	public int getReqPath() {
		return reqPath;
	}

	public void setReqPath(int reqPath) {
		this.reqPath = reqPath;
	}

	public int getStep() {
		return step;
	}

	public void setStep(int step) {
		this.step = step;
	}

	public Object getObject() {
		return object;
	}

	public void setObject(Object object) {
		this.object = object;
	}

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

}
