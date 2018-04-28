package com.acmr.excel.service;



import java.util.List;

import org.springframework.stereotype.Service;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;
@Service
public class SheetService {
	/**
	 * 增加行，用于初始化时向下滚动
	 * 
	 * @param sheet
	 *            SheetElement
	 * @param rowNum
	 *            增加行数
	 */
	public void addRowLine(ExcelSheet sheet, int rowNum) {
		for (int i = 0; i < rowNum; i++) {
			sheet.addRow();
		}
	}
	/**
	 * 增加列，用于初始化时向右滚动
	 * 
	 * @param sheet
	 *            SheetElement
	 * @param rowNum
	 *            增加行数
	 */
	public void addColLine(ExcelSheet sheet, int rowNum) {
		for (int i = 0; i < rowNum; i++) {
			sheet.addColumn();
		}
	}
	/**
	 * 冻结
	 * 
	 * @param excelSheet
	 *            excelSheet对象
	 * @param frozenY
	 *            冻结横坐标
	 * @param frozenX
	 *            冻结纵坐标
	 * @param startY
	 *            开始点横坐标
	 * @param startX
	 *            开始点纵坐标
	 */
	public void frozen(ExcelSheet excelSheet,Frozen frozen) {
		ExcelSheetFreeze excelSheetFreeze = excelSheet.getFreeze();
		if (excelSheetFreeze == null) {
			excelSheetFreeze = new ExcelSheetFreeze();
			excelSheet.setFreeze(excelSheetFreeze);
		}
		excelSheet.getExps().remove("fr");
		excelSheet.getExps().remove("fc");
		int frozenYIndex = frozen.getOprRow();
		int frozenXIndex = frozen.getOprCol();
		if(frozenYIndex == -1){
			frozenYIndex = 0;
			excelSheet.getExps().put("fr", "fr");
		}
		if(frozenXIndex == -1){
			frozenXIndex = 0;
			excelSheet.getExps().put("fc", "fc");
		}
		excelSheetFreeze.setRow(frozenYIndex);
		excelSheetFreeze.setCol(frozenXIndex);
		excelSheetFreeze.setFirstrow(frozenYIndex);
		excelSheetFreeze.setFirstcol(frozenXIndex);
	}
	
	public void cancelColHide(ExcelSheet excelSheet,ColOperate colOperate) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for(ExcelColumn excelColumn : colList){
			if(excelColumn.isColumnhidden()){
				excelColumn.setColumnhidden(false);
			}
		}
	}
	public void cancelRowHide(ExcelSheet excelSheet,RowOperate rowOperate) {
		List<ExcelRow> rowList = excelSheet.getRows();
		for(ExcelRow excelRow : rowList){
			if(excelRow.isRowhidden()){
				excelRow.setRowhidden(false);
			}
		}
	}
	
	
	
	public void undo(VersionHistory versionHistory,int step,ExcelSheet sheet){
		int version = versionHistory.getVersion().get(step-1);
		History hisory= versionHistory.getMap().get(version);
		int operatorType = hisory.getOperatorType();
		List<ChangeArea> changeAreaList = hisory.getChangeAreaList();
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		versionHistory.getVersion().put(step, version-1);
		switch (operatorType) {
		case OperatorConstant.textData:
		case OperatorConstant.merge:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontsize:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontfamily:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontweight:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontitalic:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fontcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.wordWrap:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
			}
			break;
		case OperatorConstant.fillbgcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.textDataformat:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;

		case OperatorConstant.commentset:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				if(excelCell == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					excelCell.setMemo((String)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.mergedelete:
			sheet.MergedRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		case OperatorConstant.frame:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.alignlevel:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.alignvertical:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
					excelCell.setCellstyle((ExcelCellStyle)changeArea.getOriginalValue());
				}
			}
			break;
//		case OperatorConstant.rowsinsert:
//			break;
//		case OperatorConstant.rowsdelete:
//			break;
//		case OperatorConstant.colsinsert:
//			break;
//		case OperatorConstant.colsdelete:
//			break;
		case OperatorConstant.paste:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.copy:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
		case OperatorConstant.cut:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				if(changeArea.getOriginalValue() == null){
					rowList.get(rowIndex).getCells().set(colIndex, null);
				}else{
					rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getOriginalValue());
				}
			}
			break;
//		case OperatorConstant.frozen:
//			break;
//		case OperatorConstant.unFrozen:
//			break;
//		case OperatorConstant.colswidth:
//			break;
//		case OperatorConstant.colshide:
//			break;	
//		case OperatorConstant.rowshide:
//			break;	
//		case OperatorConstant.colhideCancel:
//			break;	
//		case OperatorConstant.rowhideCancel:
//			break;	
//		case OperatorConstant.rowsheight:
//			break;
//		case OperatorConstant.addRowLine:
//			break;
//		case OperatorConstant.addColLine:
//			break;	
//		case OperatorConstant.colorset:
//		break;
		default:
			break;
		}
	}
	
	public void redo(VersionHistory versionHistory,int step,ExcelSheet sheet){
		int version = versionHistory.getVersion().get(step-1);
		History hisory= versionHistory.getMap().get(version+1);
		int operatorType = hisory.getOperatorType();
		List<ChangeArea> changeAreaList = hisory.getChangeAreaList();
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		versionHistory.getVersion().put(step, version+1);
		switch (operatorType) {
		case OperatorConstant.textData:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.fontsize:
		case OperatorConstant.fontfamily:
		case OperatorConstant.fontweight:
		case OperatorConstant.fontitalic:
		case OperatorConstant.fontcolor:
		case OperatorConstant.wordWrap:
		case OperatorConstant.alignlevel:
		case OperatorConstant.alignvertical:
		case OperatorConstant.fillbgcolor:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				if(excelCell == null){
					excelCell = new ExcelCell();
					rowList.get(rowIndex).getCells().set(colIndex, excelCell);
				}
				excelCell.setCellstyle((ExcelCellStyle)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.textDataformat:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;

		case OperatorConstant.commentset:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				ExcelCell excelCell = rowList.get(rowIndex).getCells().get(colIndex);
				excelCell.setMemo(changeArea.getUpdateValue().toString());
			}
			break;
		case OperatorConstant.merge:
			sheet.MergedRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		case OperatorConstant.mergedelete:
			sheet.SplitRegions(hisory.getMergerRowStart(), hisory.getMergerColStart(), hisory.getMergerRowEnd(),
					hisory.getMergerColEnd());
			break;
		
//		case OperatorConstant.rowsinsert:
//			break;
//		case OperatorConstant.rowsdelete:
//			break;
//		case OperatorConstant.colsinsert:
//			break;
//		case OperatorConstant.colsdelete:
//			break;
		case OperatorConstant.paste:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.copy:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
		case OperatorConstant.cut:
			for (ChangeArea changeArea : changeAreaList) {
				int colIndex = changeArea.getColIndex();
				int rowIndex = changeArea.getRowIndex();
				rowList.get(rowIndex).getCells().set(colIndex, (ExcelCell)changeArea.getUpdateValue());
			}
			break;
//		case OperatorConstant.frozen:
//			break;
//		case OperatorConstant.unFrozen:
//			break;
//		case OperatorConstant.colswidth:
//			break;
//		case OperatorConstant.colshide:
//			break;	
//		case OperatorConstant.rowshide:
//			break;	
//		case OperatorConstant.colhideCancel:
//			break;	
//		case OperatorConstant.rowhideCancel:
//			break;	
//		case OperatorConstant.rowsheight:
//			break;
//		case OperatorConstant.addRowLine:
//			break;
//		case OperatorConstant.addColLine:
//			break;	
//		case OperatorConstant.colorset:
//		break;
		default:
			break;
		}
	}
	
	
	
	
}
