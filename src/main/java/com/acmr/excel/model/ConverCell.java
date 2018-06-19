package com.acmr.excel.model;

import com.acmr.excel.model.complete.Occupy;

public class ConverCell {
	private String id;
	private Occupy occupy = new Occupy();

	public String getId() {
		return id;
	}

	public void setId(String id) {
		this.id = id;
	}

	public Occupy getOccupy() {
		return occupy;
	}

	public void setOccupy(Occupy occupy) {
		this.occupy = occupy;
	}

}
