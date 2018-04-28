package com.acmr.excel.service;



import org.springframework.stereotype.Service;

import acmr.excel.ExcelException;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.complete.GLXY;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.util.List;


/**
 * 单元格操作service
 * 
 * @author jinhr
 *
 */

@Service
public class CellService {
	/**
	 * 合并单元格
	 * 
	 * @param cell
	 *            Cell对象
	 */
	public void mergeCell(ExcelSheet excelSheet, Cell cell,VersionHistory versionHistory,int step) {
		versionHistory.getVersion().put(step, step);
		History history = new History();
		history.setOperatorType(OperatorConstant.merge);
		int firstRow = cell.getCoordinate().getStartRow();
		int firstCol = cell.getCoordinate().getStartCol();
		int lastRow = cell.getCoordinate().getEndRow();
		int lastCol = cell.getCoordinate().getEndCol();
		for (int i = firstRow; i <= lastRow; i++) {
			for (int j = firstCol; j <= lastCol; j++) {
				ChangeArea changeArea = new ChangeArea();
				changeArea.setRowIndex(i);
				changeArea.setColIndex(j);
				ExcelCell excelCell = excelSheet.getRows().get(i).getCells().get(j);
				if (excelCell == null) {
					changeArea.setOriginalValue(null);
				} else {
					changeArea.setOriginalValue(excelCell.clone());
				}
				history.getChangeAreaList().add(changeArea);
			}
		}
		history.setMergerRowStart(firstRow);
		history.setMergerRowEnd(lastRow);
		history.setMergerColStart(firstCol);
		history.setMergerColEnd(lastCol);
		versionHistory.getMap().put(step, history);
		excelSheet.MergedRegions(firstRow, firstCol, lastRow, lastCol);
		
		
	}

	/**
	 * 单元格拆分
	 * 
	 * @param excelSheet
	 *            ExcelSheet对象
	 * @param cell
	 *            Cell对象
	 */

	public void splitCell(ExcelSheet excelSheet, Cell cell,VersionHistory versionHistory,int step) {
		versionHistory.getVersion().put(step, step);
		History history = new History();
		history.setOperatorType(OperatorConstant.mergedelete);
		int firstRow = cell.getCoordinate().getStartRow();
		int firstCol = cell.getCoordinate().getStartCol();
		int lastRow = cell.getCoordinate().getEndRow();
		int lastCol = cell.getCoordinate().getEndCol();
		for (int i = firstRow; i <= lastRow; i++) {
			for (int j = firstCol; j <= lastCol; j++) {
				ChangeArea changeArea = new ChangeArea();
				changeArea.setRowIndex(i);
				changeArea.setColIndex(j);
				ExcelCell excelCell = excelSheet.getRows().get(i).getCells().get(j);
				if (excelCell == null) {
					changeArea.setOriginalValue(null);
				} else {
					changeArea.setOriginalValue(excelCell.clone());
				}
				history.getChangeAreaList().add(changeArea);
			}
		}
		history.setMergerRowStart(firstRow);
		history.setMergerRowEnd(lastRow);
		history.setMergerColStart(firstCol);
		history.setMergerColEnd(lastCol);
		excelSheet.SplitRegions(firstRow, firstCol, lastRow, lastCol);
		versionHistory.getMap().put(step, history);
	}

	/**
	 * 增加行
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void addRow(ExcelSheet excelSheet, RowOperate rowOperate) {
		excelSheet.insertRow(rowOperate.getRow());
	}

	/**
	 * 删除行
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void deleteRow(ExcelSheet excelSheet, RowOperate rowOperate) {
		excelSheet.delRow(rowOperate.getRow());
	}

	/**
	 * 增加列
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void addCol(ExcelSheet excelSheet, ColOperate colOperate) {
		excelSheet.insertColumn(colOperate.getCol());
		List<ExcelColumn> colList = excelSheet.getCols();
		excelSheet.delColumn(colList.size()-1);
	}

	/**
	 * 删除列
	 * 
	 * @param sheet
	 *            SpreadSheet对象
	 * @param cell
	 *            Cell对象
	 */
	public void deleteCol(ExcelSheet excelSheet, ColOperate colOperate) {
		excelSheet.delColumn(colOperate.getCol());
		List<ExcelColumn> excelColumns = excelSheet.getCols();
		if(excelColumns.size() < 26){
			excelSheet.addColumn();
		}
	}

	/**
	 * 调整列宽
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param colAlais
	 *            列索引
	 * @param offset
	 *            偏移量
	 */
	public void controlColWidth(ExcelSheet excelSheet,ColWidth colWidth) {
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		int colIndex = colWidth.getCol();
		int offsetWidth = colWidth.getOffset();
		colList.get(colIndex).setWidth(offsetWidth);
	}
	/**
	 * 列隐藏
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param colAlais
	 *            列索引
	 * @param offset
	 *            偏移量
	 */
	public void colHide(ExcelSheet excelSheet,ColOperate colHide) {
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		int colIndex = colHide.getCol();
		colList.get(colIndex).setColumnhidden(true);
	}
	/**
	 * 行隐藏
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param colAlais
	 *            列索引
	 * @param offset
	 *            偏移量
	 */
	public void rowHide(ExcelSheet excelSheet,RowOperate rowHide) {
		int rowIndex = rowHide.getRow();
		excelSheet.getRows().get(rowIndex).setRowhidden(true);
	}
	/**
	 * 调整行高
	 * 
	 * @param excelSheet
	 *            SpreadSheet对象
	 * @param rowAlais
	 *            行索引
	 * @param offset
	 *            偏移量
	 */
	public void controlRowHeight(ExcelSheet excelSheet, RowHeight rowHeight) {
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		int rowIndex = rowHeight.getRow();
		int offsetHeight = rowHeight.getOffset();
		rowList.get(rowIndex).setHeight(offsetHeight);
	}

}
