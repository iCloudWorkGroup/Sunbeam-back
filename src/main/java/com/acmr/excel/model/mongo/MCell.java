package com.acmr.excel.model.mongo;

import java.io.Serializable;

import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Occupy;

public class MCell implements Serializable {
	/* id */
	private String id;
	/* sheet id */
	private String sheetId;
	/* 合并几行 */
	private Integer rowspan;
	/* 合并几列 */
	private Integer colspan;
	/* 边框 */
	private Border border = new Border();
	/* 内容 */
	private Content content = new Content();
	/* 注释 */
	private CustomProp customProp = new CustomProp();
	/* 用于复制操作，相对于起始行的距离 */
	private Integer row;
	/* 用于复制操作，相对于起始行的距离 */
	private Integer col;
	
	private Occupy occupy = new Occupy();

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

	public Integer getRow() {
		return row;
	}

	public void setRow(Integer row) {
		this.row = row;
	}

	public Integer getCol() {
		return col;
	}

	public void setCol(Integer col) {
		this.col = col;
	}

	public MCell() {
          
	}

	public MCell(String id, String sheetId) {
		this.id = id;
		this.colspan = 1;
		this.rowspan = 1;
		this.sheetId = sheetId;
		this.getContent().setAlignCol("middle");
		this.getContent().setLocked(true);
		// this.getContent().setAlignRow("center");

	}

	public MCell(MCell mc) {
		this.sheetId = mc.getSheetId();
		this.content = mc.getContent().clone();
		this.border = mc.getBorder();
		this.customProp = mc.getCustomProp();
	}
	
	public MCell(Content content) {
		
		this.content = content;
		
	}

	public Occupy getOccupy() {
		return occupy;
	}

	public void setOccupy(Occupy occupy) {
		this.occupy = occupy;
	}

}
