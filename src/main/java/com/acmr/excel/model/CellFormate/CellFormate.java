package com.acmr.excel.model.CellFormate;

import java.io.Serializable;

import com.acmr.excel.model.Coordinate;

public class CellFormate implements Serializable {
	private Coordinate coordinate;
	private String format;
	private int decimalPoint;
	private boolean thousandPoint;
	private String dateFormat;
	private String currencySymbol;


	public Coordinate getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(Coordinate coordinate) {
		this.coordinate = coordinate;
	}

	public String getFormat() {
		return format;
	}

	public void setFormat(String format) {
		this.format = format;
	}

	public int getDecimalPoint() {
		return decimalPoint;
	}

	public void setDecimalPoint(int decimalPoint) {
		this.decimalPoint = decimalPoint;
	}

	public boolean isThousandPoint() {
		return thousandPoint;
	}

	public void setThousandPoint(boolean thousandPoint) {
		this.thousandPoint = thousandPoint;
	}

	public String getCurrencySymbol() {
		return currencySymbol;
	}

	public void setCurrencySymbol(String currencySymbol) {
		this.currencySymbol = currencySymbol;
	}

	public String getDateFormat() {
		return dateFormat;
	}

	public void setDateFormat(String dateFormat) {
		this.dateFormat = dateFormat;
	}

}
