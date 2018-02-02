package com.acmr.excel.model.complete;

public class ReturnParam {
	private String displayRowStartAlias;
	private String displayColStartAlias;
	private int dataRowStartIndex;
	private int dataColStartIndex;
	private int maxPixel;
	private int maxRowPixel;
	private int maxColPixel;

	public int getMaxRowPixel() {
		return maxRowPixel;
	}

	public void setMaxRowPixel(int maxRowPixel) {
		this.maxRowPixel = maxRowPixel;
	}

	public int getMaxColPixel() {
		return maxColPixel;
	}

	public void setMaxColPixel(int maxColPixel) {
		this.maxColPixel = maxColPixel;
	}

	public String getDisplayRowStartAlias() {
		return displayRowStartAlias;
	}

	public void setDisplayRowStartAlias(String displayRowStartAlias) {
		this.displayRowStartAlias = displayRowStartAlias;
	}

	public String getDisplayColStartAlias() {
		return displayColStartAlias;
	}

	public void setDisplayColStartAlias(String displayColStartAlias) {
		this.displayColStartAlias = displayColStartAlias;
	}

	public int getDataRowStartIndex() {
		return dataRowStartIndex;
	}

	public void setDataRowStartIndex(int dataRowStartIndex) {
		this.dataRowStartIndex = dataRowStartIndex;
	}

	public int getDataColStartIndex() {
		return dataColStartIndex;
	}

	public void setDataColStartIndex(int dataColStartIndex) {
		this.dataColStartIndex = dataColStartIndex;
	}

	public int getMaxPixel() {
		return maxPixel;
	}

	public void setMaxPixel(int maxPixel) {
		this.maxPixel = maxPixel;
	}

}
