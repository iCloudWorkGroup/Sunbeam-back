package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.List;

import com.acmr.excel.model.mongo.MRowColList;

public class Before implements Serializable{

	private MColBefore mcol;
	
	private MRowBefore mrow;
	
	private MCellBefore mcell;
	
	private MRowColCellBefore mrowColCell;
	//简化行列表
	private List<MRowColList> rowList;
	//简化列列表
	private List<MRowColList> colList;
	
	private List<Object> delList;
	//标识，这步操作是否可以执行  1 可以  2 可以
	private int sure = 1;

	public MColBefore getMcol() {
		return mcol;
	}

	public void setMcol(MColBefore mcol) {
		this.mcol = mcol;
	}

	public MRowBefore getMrow() {
		return mrow;
	}

	public void setMrow(MRowBefore mrow) {
		this.mrow = mrow;
	}

	public MCellBefore getMcell() {
		return mcell;
	}

	public void setMcell(MCellBefore mcell) {
		this.mcell = mcell;
	}

	public MRowColCellBefore getMrowColCell() {
		return mrowColCell;
	}

	public void setMrowColCell(MRowColCellBefore mrowColCell) {
		this.mrowColCell = mrowColCell;
	}

	public List<MRowColList> getRowList() {
		return rowList;
	}

	public void setRowList(List<MRowColList> rowList) {
		this.rowList = rowList;
	}

	public List<MRowColList> getColList() {
		return colList;
	}

	public void setColList(List<MRowColList> colList) {
		this.colList = colList;
	}

	public List<Object> getDelList() {
		return delList;
	}

	public void setDelList(List<Object> delList) {
		this.delList = delList;
	}

	public int getSure() {
		return sure;
	}

	public void setSure(int sure) {
		this.sure = sure;
	}
	
}
