package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.List;

import com.acmr.excel.model.mongo.MRow;

public class MRowBefore implements Serializable{
  
	private List<String> idList;
	
	private List<MRow> mrowList;

	public List<String> getIdList() {
		return idList;
	}

	public void setIdList(List<String> idList) {
		this.idList = idList;
	}

	public List<MRow> getMrowList() {
		return mrowList;
	}

	public void setMrowList(List<MRow> mrowList) {
		this.mrowList = mrowList;
	}
	
	public MRowBefore(List<String> idList,List<MRow> mrowList){
		this.idList = idList;
		this.mrowList = mrowList;
	}
	
}
