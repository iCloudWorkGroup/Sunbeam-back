package com.acmr.excel.model.history;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Record implements Serializable{
	// 存放修改的对象
	private List<Param> param = new ArrayList<Param>();
	
	//标识，与这步操作是否可以执行  1 可以  2 不可以
	private int sure = 2;

	public List<Param> getParam() {
		return param;
	}

	public void setParam(List<Param> param) {
		this.param = param;
	}

	public int getSure() {
		return sure;
	}

	public void setSure(int sure) {
		this.sure = sure;
	}
}
