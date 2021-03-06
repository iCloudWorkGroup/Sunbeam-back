package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.HashMap;
import java.util.Map;

public class Content implements Serializable, Cloneable {
	private String alignCol;
	private String alignRow;
	private String alignRowFormat;
	private Boolean weight;
	private String color;
	private String displayTexts;
	private String family;
	private Boolean italic;
	private String size;
	private String texts;
	private Integer underline;
	private Boolean wordWrap;//自动换行
	private String type;
	private String express;
	private String background;
	private Boolean locked ;

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

	public Boolean getWeight() {
		return weight;
	}

	public void setWeight(Boolean weight) {
		this.weight = weight;
	}

	public Boolean getWordWrap() {
		return wordWrap;
	}

	public void setWordWrap(Boolean wordWrap) {
		this.wordWrap = wordWrap;
	}

	public Boolean getLocked() {
		return locked;
	}

	public void setLocked(Boolean locked) {
		this.locked = locked;
	}

	public Integer getUnderline() {
		return underline;
	}

	public void setUnderline(Integer underline) {
		this.underline = underline;
	}

	public String getType() {
		return type;
	}

	public void setType(String type) {
		this.type = type;
	}

	public String getExpress() {
		return express;
	}

	public void setExpress(String express) {
		this.express = express;
	}

	public String getBackground() {
		return background;
	}

	public void setBackground(String background) {
		this.background = background;
	}

	public Content clone() {
		Content o = null;
		try {
			o = (Content) super.clone();
		} catch (CloneNotSupportedException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return o;
	}

	public String getAlignRowFormat() {
		return alignRowFormat;
	}

	public void setAlignRowFormat(String alignRowFormat) {
		this.alignRowFormat = alignRowFormat;
	}

	public static void main(String[] args) {
		Map<String, String> map = new HashMap<String, String>();
		map.put("1", "1");
		System.out.println(map.remove("2"));
	}
}
