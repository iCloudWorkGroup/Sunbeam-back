package com.acmr.excel.model.mongo;

import java.io.Serializable;

import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;

public class MCell implements Serializable{
	/*id*/
	private String id;
	/*sheet id*/
	private String sheetId;
	/*合并几行*/
	private Integer rowspan;
	/*合并几列*/
	private Integer colspan;
	/*边框*/
	private Border border = new Border();
	/*内容*/
	private Content content = new Content();
	/*注释*/
	private CustomProp customProp = new CustomProp();
	public String getId() {
		return id;
	}
	public void setId(String id) {
		this.id = id;
	}
	public String getSheetId() {
		return sheetId;
	}
	public void setSheetId(String sheetId) {
		this.sheetId = sheetId;
	}
	public Integer getRowspan() {
		return rowspan;
	}
	public void setRowspan(Integer rowspan) {
		this.rowspan = rowspan;
	}
	public Integer getColspan() {
		return colspan;
	}
	public void setColspan(Integer colspan) {
		this.colspan = colspan;
	}
	public Border getBorder() {
		return border;
	}
	public void setBorder(Border border) {
		this.border = border;
	}
	public Content getContent() {
		return content;
	}
	public void setContent(Content content) {
		this.content = content;
	}
	public CustomProp getCustomProp() {
		return customProp;
	}
	public void setCustomProp(CustomProp customProp) {
		this.customProp = customProp;
	}

}
