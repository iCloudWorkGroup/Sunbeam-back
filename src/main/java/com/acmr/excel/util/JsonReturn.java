package com.acmr.excel.util;


/**
 * 自定义返回json类
 *
 */
public class JsonReturn {
	private static final long serialVersionUID = 1L;
	private int returncode;
	private Object returndata;
	private Object param1;
	private Object param2;
	private Object param3;
	private Object param4;
	private Object param5;
	/**
	 * 行号
	 */
	private int rowNum;
	/**
	 * 列号
	 */
	private int colNum;

	/**
	 * 总像素
	 */
	private int rowLength;
	/**
	 * 
	 * @return
	 */
	private int maxPixel;

	private String displayRowStartAlias;
	private String displayColStartAlias;
	private int dataRowStartIndex;
	private int dataColStartIndex;
	private int maxRowPixel;
	private int maxColPixel;
	private String aliasRowCounter;
	private String aliasColCounter;

	// private Object returnParam;
	
	

	public int getMaxPixel() {
		return maxPixel;
	}

	public String getAliasRowCounter() {
		return aliasRowCounter;
	}

	public void setAliasRowCounter(String aliasRowCounter) {
		this.aliasRowCounter = aliasRowCounter;
	}

	public String getAliasColCounter() {
		return aliasColCounter;
	}

	public void setAliasColCounter(String aliasColCounter) {
		this.aliasColCounter = aliasColCounter;
	}

	public int getReturncode() {
		return returncode;
	}

	public void setReturncode(int returncode) {
		this.returncode = returncode;
	}

	public Object getReturndata() {
		return returndata;
	}

	public void setReturndata(Object returndata) {
		this.returndata = returndata;
	}

	public Object getParam1() {
		return param1;
	}

	public void setParam1(Object param1) {
		this.param1 = param1;
	}

	public Object getParam2() {
		return param2;
	}

	public void setParam2(Object param2) {
		this.param2 = param2;
	}

	public Object getParam3() {
		return param3;
	}

	public void setParam3(Object param3) {
		this.param3 = param3;
	}

	public Object getParam4() {
		return param4;
	}

	public void setParam4(Object param4) {
		this.param4 = param4;
	}

	public Object getParam5() {
		return param5;
	}

	public void setParam5(Object param5) {
		this.param5 = param5;
	}

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

	public void setMaxPixel(int maxPixel) {
		this.maxPixel = maxPixel;
	}

	public int getRowLength() {
		return rowLength;
	}

	public void setRowLength(int rowLength) {
		this.rowLength = rowLength;
	}

	public int getRowNum() {
		return rowNum;
	}

	public void setRowNum(int rowNum) {
		this.rowNum = rowNum;
	}

	public int getColNum() {
		return colNum;
	}

	public void setColNum(int colNum) {
		this.colNum = colNum;
	}

	public JsonReturn(int code1, String msg1) {
		this.returncode = code1;
		this.returndata = msg1;
	}

	public JsonReturn(Object data1) {
		this.returncode = 200;
		this.returndata = data1;
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

}
