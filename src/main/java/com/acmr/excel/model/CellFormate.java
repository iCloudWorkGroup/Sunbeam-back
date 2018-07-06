package com.acmr.excel.model;

import java.io.Serializable;
import java.util.List;

public class CellFormate implements Serializable {
	private List<Coordinate> coordinate;
	private String format;
	private String express;
	public List<Coordinate> getCoordinate() {
		return coordinate;
	}
	public void setCoordinate(List<Coordinate> coordinate) {
		this.coordinate = coordinate;
	}
	public String getFormat() {
		return format;
	}
	public void setFormat(String format) {
		this.format = format;
	}
	public String getExpress() {
		return express;
	}
	public void setExpress(String express) {
		this.express = express;
	}

}
