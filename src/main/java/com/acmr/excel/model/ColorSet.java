package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class ColorSet implements Serializable {
	private List<Coordinate> coordinates;
	private String color;

	public List<Coordinate> getCoordinates() {
		return coordinates;
	}

	public void setCoordinates(List<Coordinate> coordinates) {
		this.coordinates = coordinates;
	}

	public String getColor() {
		return color;
	}

	public void setColor(String color) {
		this.color = color;
	}

}
