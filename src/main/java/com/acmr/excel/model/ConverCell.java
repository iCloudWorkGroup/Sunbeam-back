package com.acmr.excel.model;

import java.io.Serializable;

import com.acmr.excel.model.complete.Occupy;

public class ConverCell implements Serializable{
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
