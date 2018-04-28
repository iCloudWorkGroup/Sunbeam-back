package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class History implements Serializable{
	private Integer OperatorType;
	private List<ChangeArea> changeAreaList = new ArrayList<ChangeArea>();
	private Integer mergerColStart;
	private Integer mergerColEnd;
	private Integer mergerRowStart;
	private Integer mergerRowEnd;
	public Integer getOperatorType() {
		return OperatorType;
	}
	public void setOperatorType(Integer operatorType) {
		OperatorType = operatorType;
	}
	public List<ChangeArea> getChangeAreaList() {
		return changeAreaList;
	}
	public void setChangeAreaList(List<ChangeArea> changeAreaList) {
		this.changeAreaList = changeAreaList;
	}
	public Integer getMergerColStart() {
		return mergerColStart;
	}
	public void setMergerColStart(Integer mergerColStart) {
		this.mergerColStart = mergerColStart;
	}
	public Integer getMergerColEnd() {
		return mergerColEnd;
	}
	public void setMergerColEnd(Integer mergerColEnd) {
		this.mergerColEnd = mergerColEnd;
	}
	public Integer getMergerRowStart() {
		return mergerRowStart;
	}
	public void setMergerRowStart(Integer mergerRowStart) {
		this.mergerRowStart = mergerRowStart;
	}
	public Integer getMergerRowEnd() {
		return mergerRowEnd;
	}
	public void setMergerRowEnd(Integer mergerRowEnd) {
		this.mergerRowEnd = mergerRowEnd;
	}

	
}
