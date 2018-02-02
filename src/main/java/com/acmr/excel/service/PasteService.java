package com.acmr.excel.service;



import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.util.ListHashMap;

import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.OuterPasteData;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.copy.TempObj;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Service;
@Service
public class PasteService {
	/**
	 * 是否可以复制
	 * 
	 * @param copy
	 * @param excelBook
	 * @return
	 */
	public boolean canCopy(Copy copy, ExcelBook excelBook) {
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int targetRowIndex = copy.getTarget().getOprRow();
		int targetColIndex = copy.getTarget().getOprCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		for (int i = targetRowIndex; i < targetRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = targetColIndex; j < targetColIndex + colRange; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}

	/**
	 * 粘贴
	 * 
	 * @param paste
	 *            paste对象
	 * @param excelBook
	 *            excelBook对象
	 */

	public void data(Paste paste, ExcelBook excelBook,VersionHistory versionHistory,int step) {
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
		}
		version += 1;
		versionHistory.getVersion().put(step, version);
		History history = new History();
		history.setOperatorType(OperatorConstant.paste);
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int colLength = excelSheet.getCols().size();
		int colRange = paste.getColLen();
		if (colLength < paste.getOprCol() + colRange) {
			for (int i = 0; i < paste.getOprCol() + colRange - colLength; i++) {
				excelSheet.addColumn();
			}
		}
		List<OuterPasteData> pasteList = paste.getPasteData();
		for (OuterPasteData outerPasteData : pasteList) {
			int rowIndex = outerPasteData.getRow();
			ExcelRow excelRow = rowList.get(rowIndex);
			int colIndex = outerPasteData.getCol();
			List<ExcelCell> cellList = excelRow.getCells();
			ExcelCell excelCell = cellList.get(colIndex);
			ChangeArea changeArea = new ChangeArea();
			changeArea.setColIndex(colIndex);
			changeArea.setRowIndex(rowIndex);
			if (excelCell == null) {
				changeArea.setOriginalValue(null);
//				excelCell = new ExcelCell();
//				excelCell.getCellstyle().setBgcolor(new ExcelColor(255, 255, 255));
//				cellList.set(colIndex, excelCell);
			}else{
				changeArea.setOriginalValue(excelCell.clone());
			}
//			String text = outerPasteData.getContent();
//				//text = java.net.URLDecoder.decode(text, "utf-8");
//			excelCell.setText(text);
//			excelCell.setType(CELLTYPE.STRING);
//			excelCell.setValue(text);
			changeArea.setUpdateValue(excelCell);
			history.getChangeAreaList().add(changeArea);
		}
		versionHistory.getMap().put(version, history);
		int row = paste.getOprRow();
		int col = paste.getOprCol();
		for (int i = row; i < row + paste.getRowLen(); i++) {
			for (int j = col; j < col + paste.getColLen(); j++) {
				rowList.get(i).set(j, null);
			}
		}
		for (OuterPasteData outerPasteData : pasteList) {
			int rowIndex = outerPasteData.getRow();
			ExcelRow excelRow = rowList.get(rowIndex);
			int colIndex = outerPasteData.getCol();
			List<ExcelCell> cellList = excelRow.getCells();
			ExcelCell excelCell = cellList.get(colIndex);
			if (excelCell == null) {
				excelCell = new ExcelCell();
				excelCell.getCellstyle().setBgcolor(new ExcelColor(255, 255, 255));
				cellList.set(colIndex, excelCell);
			}
			String text = outerPasteData.getContent();
//				//text = java.net.URLDecoder.decode(text, "utf-8");
			excelCell.setText(text);
			excelCell.setType(CELLTYPE.STRING);
			excelCell.setValue(text);
		}
	}

	/**
	 * 复制粘贴
	 * 
	 * @param copy
	 * @param excelBook
	 */
	public void copy(Copy copy, ExcelBook excelBook,VersionHistory versionHistory,int step) {
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
		}
		version += 1;
		versionHistory.getVersion().put(step, version);
		History history = new History();
		history.setOperatorType(OperatorConstant.copy);
		copyOrCut(copy, excelBook, null,history);
		versionHistory.getMap().put(version, history);
	}

	/**
	 * 剪切粘贴
	 * 
	 * @param copy
	 * @param excelBook
	 */

	public void cut(Copy copy, ExcelBook excelBook,VersionHistory versionHistory,int step) {
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
		}
		version += 1;
		versionHistory.getVersion().put(step, version);
		History history = new History();
		history.setOperatorType(OperatorConstant.cut);
		copyOrCut(copy, excelBook, "cut",history);
		versionHistory.getMap().put(version, history);
	}

	/**
	 * 复制或剪切
	 * 
	 * @param copy
	 * @param excelBook
	 * @param flag
	 *            copy:复制 cut:剪切
	 */

	private void copyOrCut(Copy copy, ExcelBook excelBook, String flag,History history) {
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int targetRowIndex = copy.getTarget().getOprRow();
		int targetColIndex = copy.getTarget().getOprCol();
		int colLength = excelSheet.getCols().size();
		int colRange = endColIndex - startColIndex;
		if (colLength < targetColIndex + colRange) {
			for (int i = 0; i <= targetColIndex + colRange - colLength; i++) {
				excelSheet.addColumn();
			}
		}
		List<TempObj> temList = new ArrayList<TempObj>();
		for (int i = startRowIndex; i <= endRowIndex; i++) {
			ExcelRow excelRow = rowList.get(i);
			int tempColIndex = targetColIndex;
			for (int j = startColIndex; j <= endColIndex; j++) {
				ExcelCell excelCell = excelRow.getCells().get(j);
				ChangeArea changeArea = new ChangeArea();
				changeArea.setColIndex(tempColIndex);
				changeArea.setRowIndex(targetRowIndex);
				if(excelCell == null){
					changeArea.setOriginalValue(null);
					excelCell = new ExcelCell();
				}else{
					changeArea.setOriginalValue(rowList.get(targetRowIndex).getCells().get(tempColIndex));
				}
				ExcelCell newExcelCell = excelCell.clone();
				//rowList.get(targetRowIndex).set(tempColIndex, newExcelCell);
				TempObj tempObj = new TempObj();
				tempObj.setRow(targetRowIndex);
				tempObj.setCol(tempColIndex);
				tempObj.setExcelCell(newExcelCell);
				temList.add(tempObj);
				changeArea.setUpdateValue(newExcelCell);
				history.getChangeAreaList().add(changeArea);
				if ("cut".equals(flag)) {
					rowList.get(i).set(j, null);
					ChangeArea changeArea2 = new ChangeArea();
					changeArea2.setColIndex(i);
					changeArea2.setRowIndex(j);
					changeArea2.setOriginalValue(excelCell.clone());
					changeArea2.setUpdateValue(null);
					history.getChangeAreaList().add(changeArea2);
				}
				tempColIndex++;
			}
			targetRowIndex++;
		}

		for (TempObj tempObj : temList) {
			rowList.get(tempObj.getRow()).set(tempObj.getCol(),tempObj.getExcelCell());
		}
		
		
	}

	/**
	 * 是否可以粘贴
	 * 
	 * @param isAblePaste
	 * @param excelBook
	 * @return
	 */

	public boolean isAblePaste(Paste paste, ExcelBook excelBook) {
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = paste.getOprRow();
		int startColIndex = paste.getOprCol();
		int rowRange = paste.getRowLen();
		int colRange = paste.getColLen();
		for (int i = startRowIndex; i < startRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = startColIndex; j < startColIndex + colRange; j++) {
				if (j > excelSheet.getCols().size() - 1) {
					break;
				}
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}
	
	public boolean isCopyPaste(Copy copy, ExcelBook excelBook){
		boolean canPaste = true;
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getOprRow();
		int targetStartColIndex = copy.getTarget().getOprCol();
		int colLength = excelSheet.getCols().size();
//		if (colLength < targetStartColIndex + colRange) {
//			for (int i = 0; i <= targetStartColIndex + colRange - colLength+1; i++) {
//				excelSheet.addColumn();
//			}
//		}
		for (int i = targetStartRowIndex; i <= targetStartRowIndex + rowRange; i++) {
			ExcelRow excelRow = rowList.get(i);
			for (int j = targetStartColIndex; j <= targetStartColIndex + colRange; j++) {
				if (j > excelSheet.getCols().size() - 1) {
					break;
				}
				ExcelCell excelCell = excelRow.getCells().get(j);
				if (excelCell == null) {
					continue;
				}
				int colspan = excelCell.getColspan();
				int rowspan = excelCell.getRowspan();
				if (colspan != 1 || rowspan != 1) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;
	}
}
