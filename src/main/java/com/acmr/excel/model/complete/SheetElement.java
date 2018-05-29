package com.acmr.excel.model.complete;

import java.io.Serializable;
import java.util.ArrayList;
import java.util.List;



public class SheetElement implements Serializable{
	
	private String name = "new sheet";
	private String maxColAlias;
	private String maxRowAlias;
	private Integer maxColPixel;
	private Integer maxRowPixel;
	private String viewRowAlias;
	private String viewColAlias;
	private List<OneCell> cells = new ArrayList<OneCell>();
	private List<Gly> gridLineRow = new ArrayList<Gly>();
	private List<Glx> gridLineCol = new ArrayList<Glx>();
	private Integer sort = 0;
	private boolean protect = false;
	private List<Validate> validate = new ArrayList<Validate>();
	private Frozen frozen;
	public String getName() {
		return name;
	}
	public void setName(String name) {
		this.name = name;
	}
	public String getMaxColAlias() {
		return maxColAlias;
	}
	public void setMaxColAlias(String maxColAlias) {
		this.maxColAlias = maxColAlias;
	}
	public String getMaxRowAlias() {
		return maxRowAlias;
	}
	public void setMaxRowAlias(String maxRowAlias) {
		this.maxRowAlias = maxRowAlias;
	}
	public Integer getMaxColPixel() {
		return maxColPixel;
	}
	public void setMaxColPixel(Integer maxColPixel) {
		this.maxColPixel = maxColPixel;
	}
	public Integer getMaxRowPixel() {
		return maxRowPixel;
	}
	public void setMaxRowPixel(Integer maxRowPixel) {
		this.maxRowPixel = maxRowPixel;
	}
	public String getViewRowAlias() {
		return viewRowAlias;
	}
	public void setViewRowAlias(String viewRowAlias) {
		this.viewRowAlias = viewRowAlias;
	}
	public String getViewColAlias() {
		return viewColAlias;
	}
	public void setViewColAlias(String viewColAlias) {
		this.viewColAlias = viewColAlias;
	}
	public List<OneCell> getCells() {
		return cells;
	}
	public void setCells(List<OneCell> cells) {
		this.cells = cells;
	}
	public List<Gly> getGridLineRow() {
		return gridLineRow;
	}
	public void setGridLineRow(List<Gly> gridLineRow) {
		this.gridLineRow = gridLineRow;
	}
	public List<Glx> getGridLineCol() {
		return gridLineCol;
	}
	public void setGridLineCol(List<Glx> gridLineCol) {
		this.gridLineCol = gridLineCol;
	}
	public Integer getSort() {
		return sort;
	}
	public void setSort(Integer sort) {
		this.sort = sort;
	}
	public boolean isProtect() {
		return protect;
	}
	public void setProtect(boolean protect) {
		this.protect = protect;
	}
	public List<Validate> getValidate() {
		return validate;
	}
	public void setValidate(List<Validate> validate) {
		this.validate = validate;
	}
	public Frozen getFrozen() {
		return frozen;
	}
	public void setFrozen(Frozen frozen) {
		this.frozen = frozen;
	}

}
