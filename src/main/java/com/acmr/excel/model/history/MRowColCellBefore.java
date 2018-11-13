package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

import com.acmr.excel.model.mongo.MRowColCell;

public class MRowColCellBefore implements Serializable{

	private List<String> idList = new ArrayList<String>();
	
	private List<MRowColCell> mrowColCellList = new ArrayList<MRowColCell>();

	public List<String> getIdList() {
		return idList;
	}

	public void setIdList(List<String> idList) {
		this.idList = idList;
	}

	public List<MRowColCell> getMrowColCellList() {
		return mrowColCellList;
	}

	public void setMrowColCellList(List<MRowColCell> mrowColCellList) {
		this.mrowColCellList = mrowColCellList;
	}
	
	public MRowColCellBefore(List<String> idList,List<MRowColCell> mrowColCellList){
		this.idList = idList;
		this.mrowColCellList = mrowColCellList;
	}
	
	public MRowColCellBefore(){
		
	}
	
}
