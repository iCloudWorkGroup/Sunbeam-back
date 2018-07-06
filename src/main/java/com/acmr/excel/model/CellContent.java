package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class CellContent implements Serializable {

	private Coordinate coordinate;

	private String content;
	
	private List<RowHeight> effect;
 
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

	public List<RowHeight> getEffect() {
		return effect;
	}

	public void setEffect(List<RowHeight> effect) {
		this.effect = effect;
	}

}
