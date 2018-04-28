package com.acmr.excel.model;

import java.io.Serializable;

public class Cell implements Serializable {
	private int sheetId;
	private Coordinate coordinate = new Coordinate();
	private String direction;
	private String align;
	private String isBold;
	private String italic;
	private String wordWrap;

	/**
	 * 行索引
	 */
	private int rowIndex;
	/**
	 * 列索引
	 */
	private int colIndex;
	/**
	 * 列宽
	 */
	private int width;

	private int backIndex;

	/**
	 * 行高度
	 */
	private int height;
	/**
	 * "size": "字体大小12px",
	 */
	private String size;
	/**
	 * "family": "字体风格宋体...",
	 */
	private String family;
	/**
	 * "weight": "字体加粗",
	 */
	private boolean weight;
	/**
	 * "style": "字体倾斜italic",
	 */
	private boolean style;
	/**
	 * "color": "颜色代码（rgb）",
	 */
	private String color;
	private String bgcolor;
	/**
	 * "content" : "输入内容"
	 */
	private String content;
	/**
	 * "format" : "num: 数字类型,time：时间,text：文本"
	 */
	private String format;

	public int getSheetId() {
		return sheetId;
	}

	public void setSheetId(int sheetId) {
		this.sheetId = sheetId;
	}

	public Coordinate getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(Coordinate coordinate) {
		this.coordinate = coordinate;
	}

	public String getDirection() {
		return direction;
	}

	public void setDirection(String direction) {
		this.direction = direction;
	}

	public String getAlign() {
		return align;
	}

	public void setAlign(String align) {
		this.align = align;
	}

	public int getRowIndex() {
		return rowIndex;
	}

	public void setRowIndex(int rowIndex) {
		this.rowIndex = rowIndex;
	}

	public int getColIndex() {
		return colIndex;
	}

	public void setColIndex(int colIndex) {
		this.colIndex = colIndex;
	}

	public int getWidth() {
		return width;
	}

	public void setWidth(int width) {
		this.width = width;
	}

	public int getHeight() {
		return height;
	}

	public void setHeight(int height) {
		this.height = height;
	}

	public void setBackIndex(int backIndex) {
		this.backIndex = backIndex;
	}

	public String getSize() {
		return size;
	}

	public void setSize(String size) {
		this.size = size;
	}

	public String getFamily() {
		return family;
	}

	public void setFamily(String family) {
		this.family = family;
	}

	public boolean isWeight() {
		return weight;
	}

	public void setWeight(boolean weight) {
		this.weight = weight;
	}

	public boolean getStyle() {
		return style;
	}

	public void setStyle(boolean style) {
		this.style = style;
	}

	public String getColor() {
		return color;
	}

	public void setColor(String color) {
		this.color = color;
	}

	public String getContent() {
		return content;
	}

	public void setContent(String content) {
		this.content = content;
	}

	public String getFormat() {
		return format;
	}

	public void setFormat(String format) {
		this.format = format;
	}

	public int getBackIndex() {
		return backIndex;
	}

	public String getIsBold() {
		return isBold;
	}

	public void setIsBold(String isBold) {
		this.isBold = isBold;
	}

	public String getItalic() {
		return italic;
	}

	public void setItalic(String italic) {
		this.italic = italic;
	}

	public String getBgcolor() {
		return bgcolor;
	}

	public void setBgcolor(String bgcolor) {
		this.bgcolor = bgcolor;
	}

	public String getWordWrap() {
		return wordWrap;
	}

	public void setWordWrap(String wordWrap) {
		this.wordWrap = wordWrap;
	}

}
