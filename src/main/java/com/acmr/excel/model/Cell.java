package com.acmr.excel.model;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;

public class Cell implements Serializable {
	private int sheetId;
	private List<Coordinate> coordinate = new ArrayList<Coordinate>();
	/*边框类型 上下左右 all 、outer、none*/
	private String direction;
	/* 0 无边框 1 细边框 2 粗边框*/
	private int line;
	
	private String align;
	private String isBold;
	private Boolean italic;
	private Boolean auto;
	private List<RowHeight> effect;
	/*下划线  0 无下划线  1 有下滑线*/
	private int underline;
	//单元格锁定
	private boolean lock;
	


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
	
	private String comment;
	

	public int getSheetId() {
		return sheetId;
	}

	public void setSheetId(int sheetId) {
		this.sheetId = sheetId;
	}

	public List<Coordinate> getCoordinate() {
		return coordinate;
	}

	public void setCoordinate(List<Coordinate> coordinate) {
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

	public Boolean getItalic() {
		return italic;
	}

	public void setItalic(Boolean italic) {
		this.italic = italic;
	}

	public String getBgcolor() {
		return bgcolor;
	}

	public void setBgcolor(String bgcolor) {
		this.bgcolor = bgcolor;
	}

	public Boolean getAuto() {
		return auto;
	}

	public void setAuto(Boolean auto) {
		this.auto = auto;
	}

	public int getLine() {
		return line;
	}

	public void setLine(int line) {
		this.line = line;
	}

	public List<RowHeight> getEffect() {
		return effect;
	}

	public void setEffect(List<RowHeight> effect) {
		this.effect = effect;
	}

	public int getUnderline() {
		return underline;
	}

	public void setUnderline(int underline) {
		this.underline = underline;
	}

	public String getComment() {
		return comment;
	}

	public void setComment(String comment) {
		this.comment = comment;
	}

	public boolean isLock() {
		return lock;
	}

	public void setLock(boolean lock) {
		this.lock = lock;
	}

}
