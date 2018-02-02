package com.acmr.excel.model.complete;

import java.io.Serializable;

public class PresetCell extends BaseCell implements Serializable{
	private Long time;

	public Long getTime() {
		return time;
	}

	public void setTime(Long time) {
		this.time = time;
	}

}
