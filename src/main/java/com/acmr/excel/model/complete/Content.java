package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

public class Content implements Serializable {

	private String size = "11px";
	private String family = "SimSun";
	private Boolean bd = false;
	private Boolean italic = false;
	private String color = "rgb(0,0,0)";
	private String rgbColor = "0,0,0";
	private String alignRow = "";
	private String alignCol = "";
	private String alignLine = "bottom";
	private Integer weight;
	private String texts = "";
	private String displayTexts;

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

	public Boolean getBd() {
		return bd;
	}

	public void setBd(Boolean bd) {
		this.bd = bd;
	}

	public Boolean getItalic() {
		return italic;
	}

	public void setItalic(Boolean italic) {
		this.italic = italic;
	}

	public String getColor() {
		return color;
	}

	public void setColor(String color) {
		this.color = color;
	}

	public String getRgbColor() {
		return rgbColor;
	}

	public void setRgbColor(String rgbColor) {
		this.rgbColor = rgbColor;
	}

	public String getAlignRow() {
		return alignRow;
	}

	public void setAlignRow(String alignRow) {
		this.alignRow = alignRow;
	}

	public String getAlignCol() {
		return alignCol;
	}

	public void setAlignCol(String alignCol) {
		this.alignCol = alignCol;
	}

	public String getAlignLine() {
		return alignLine;
	}

	public void setAlignLine(String alignLine) {
		this.alignLine = alignLine;
	}

	public Integer getWeight() {
		return weight;
	}

	public void setWeight(Integer weight) {
		this.weight = weight;
	}

	public String getTexts() {
		return texts;
	}

	public void setTexts(String texts) {
		this.texts = texts;
	}

	public String getDisplayTexts() {
		return displayTexts;
	}

	public void setDisplayTexts(String displayTexts) {
		this.displayTexts = displayTexts;
	}

	public static void main(String[] args) {
		Map<String, String> map = new HashMap<String, String>();
		map.put("1", "1");
		System.out.println(map.remove("2"));
	}
}
