package com.acmr.excel.model.comment;

import java.io.Serializable;

import com.acmr.excel.model.Coordinate;


public class Comment implements Serializable{
	private String excelId;
	private Coordinate coordinate;
	private String comment;

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public Coordinate getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(Coordinate coordinate) {
		this.coordinate = coordinate;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

}
