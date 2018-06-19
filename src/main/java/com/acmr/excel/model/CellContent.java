package com.acmr.excel.model;

import java.io.Serializable;

public class CellContent implements Serializable{
	
	private Coordinate coordinate;
	
	private String content;

	public Coordinate getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(Coordinate coordinate) {
		this.coordinate = coordinate;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

}
