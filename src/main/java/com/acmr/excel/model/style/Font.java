package com.acmr.excel.model.style;

import acmr.excel.pojo.ExcelColor;

public class Font {

	private String fontname; // 字体名字
	private short size;// 字体字号 实际字号*20
	private short boldweight; // 字体黑色浓度
	private ExcelColor color; // 字体颜色
	private boolean strikeout; // 删除线
	private byte underline; // 下划线
	private boolean italic; // 斜体
	private short typeoffset; // 角标类型
	public String getFontname() {
		return fontname;
	}
	public void setFontname(String fontname) {
		this.fontname = fontname;
	}
	public short getSize() {
		return size;
	}
	public void setSize(short size) {
		this.size = size;
	}
	public short getBoldweight() {
		return boldweight;
	}
	public void setBoldweight(short boldweight) {
		this.boldweight = boldweight;
	}
	public ExcelColor getColor() {
		return color;
	}
	public void setColor(ExcelColor color) {
		this.color = color;
	}
	public boolean isStrikeout() {
		return strikeout;
	}
	public void setStrikeout(boolean strikeout) {
		this.strikeout = strikeout;
	}
	public byte getUnderline() {
		return underline;
	}
	public void setUnderline(byte underline) {
		this.underline = underline;
	}
	public boolean isItalic() {
		return italic;
	}
	public void setItalic(boolean italic) {
		this.italic = italic;
	}
	public short getTypeoffset() {
		return typeoffset;
	}
	public void setTypeoffset(short typeoffset) {
		this.typeoffset = typeoffset;
	}

	
}
