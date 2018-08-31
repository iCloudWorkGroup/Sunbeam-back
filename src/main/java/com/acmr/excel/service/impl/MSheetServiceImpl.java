package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.aop.Execute;
import com.acmr.excel.aop.HistoryAop;
import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MColDao;
import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OuterPaste;
import com.acmr.excel.model.OuterPasteData;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.SheetParam;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.history.Before;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.HistoryCache;
import com.acmr.excel.model.history.MCellBefore;
import com.acmr.excel.model.history.MColBefore;
import com.acmr.excel.model.history.MRowBefore;
import com.acmr.excel.model.history.MRowColCellBefore;
import com.acmr.excel.model.history.Param;
import com.acmr.excel.model.history.Record;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MRowColList;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MSheetService;
import com.acmr.redis.Redis;

@Service("msheetService")
public class MSheetServiceImpl implements MSheetService {
	@Resource
	private BaseDao baseDao;
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MSheetDao msheetDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MRowColCellDao mrowColCellDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MRowDao mrowDao;
	@Resource
	private Redis redis;
	// 用于记录上一步操作
	private String flag = "";

	@Override
	public void frozen(Frozen frozen, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);

		String oprRow = sortRList.get(frozen.getOprRow()).getAlias();
		String oprCol = sortCList.get(frozen.getOprCol()).getAlias();
		String viewRow = sortRList.get(frozen.getViewRow()).getAlias();
		String viewCol = sortCList.get(frozen.getViewCol()).getAlias();
		MSheet msheet = new MSheet();
		msheet.setFreeze(true);
		msheet.setRowAlias(oprRow);
		msheet.setColAlias(oprCol);
		msheet.setViewRowAlias(viewRow);
		msheet.setViewColAlias(viewCol);
		msheet.setStep(step);

		msheetDao.updateFrozen(msheet, excelId, sheetId);

	}

	@Override
	public void unFrozen(String excelId, Integer step) {
		String sheetId = excelId + 0;
		MSheet mexcel = new MSheet();
		mexcel.setFreeze(false);
		mexcel.setStep(step);
		mexcel.setId(excelId);
		msheetDao.updateUnFrozen(mexcel, excelId, sheetId);

	}

	@Override
	public int getStep(String excelId, String sheetId) {

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);

		return msheet.getStep();
	}
	
	@Override
	public void updateStep(String excelId, String sheetId) {
		msheetDao.updateStep(excelId, sheetId, 0);
	}

	@Override
	public void paste(OuterPaste outerpaste, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rowMap = new HashMap<String, Integer>();
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			rowMap.put(sortRList.get(i).getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortCList.size(); i++) {
			colMap.put(sortCList.get(i).getAlias(), i);
		}
		int oprRow = outerpaste.getOprRow();
		int oprCol = outerpaste.getOprCol();
		Paste paste = pastDataAnalysis(outerpaste.getContent());
		int rowLen = paste.getRowLen();
		int colLen = paste.getColLen();

		List<String> rowList = new ArrayList<String>();
		List<String> colList = new ArrayList<String>();
		// 查找目标区域，行或列不够的增加行列
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		int maxRow = msheet.getMaxrow();
		int maxCol = msheet.getMaxcol();

		if (sortRList.size() < oprRow + rowLen + 1) {
			for (int i = oprRow; i < sortRList.size(); i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
			// 增加新的行
			for (int i = 0; i < oprRow + rowLen + 1 - sortRList.size(); i++) {
				maxRow = maxRow + 1;
				String row = maxRow + "";
				MRow mrow = new MRow(row, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(row);
				rowCol.setLength(19);
				if (i == 0) {
					rowCol.setPreAlias(
							sortRList.get(sortRList.size() - 1).getAlias());
				} else {
					rowCol.setPreAlias((maxRow - 1) + "");
				}
				sortRList.add(rowCol);
				// 存入简化的行
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "rList");
				// 存入行样式
				baseDao.insert(excelId, mrow);
			}
		} else {
			for (int i = oprRow; i < oprRow + rowLen + 1; i++) {
				rowList.add(sortRList.get(i).getAlias());
			}

		}

		if (sortCList.size() < oprCol + colLen + 1) {

			for (int i = oprCol; i < sortCList.size(); i++) {
				colList.add(sortCList.get(i).getAlias());
			}

			// 增加新的列
			for (int i = 0; i < oprCol + colLen + 1 - sortCList.size(); i++) {
				maxCol = maxCol + 1;
				String col = maxCol + "";
				MCol mcol = new MCol(col, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(col);
				rowCol.setLength(71);
				if (i == 0) {
					rowCol.setPreAlias(
							sortCList.get(sortCList.size() - 1).getAlias());
				} else {
					rowCol.setPreAlias((maxCol - 1) + "");
				}
				sortCList.add(rowCol);
				// 存入简化的列
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
				// 存入列样式
				baseDao.insert(excelId, mcol);
			}
		} else {
			for (int i = oprCol; i < oprCol + colLen + 1; i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		}
		msheet.setMaxcol(maxCol);
		msheet.setMaxrow(maxRow);
		baseDao.update(excelId, msheet);// 更新最大行，最大列

		// 拆分目标区的合并单元格
		List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relation) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		Map<String, MCell> mcellMap = new HashMap<String, MCell>();
		cellIdList.clear();// 清除用于存贮合并单元格id
		List<String> unspanCellIdList = new ArrayList<String>();// 存储非合并单元格id
		List<Object> tempList = new ArrayList<Object>();// 用于存贮新生成的MCell和关系对象
		for (MCell mc : mcellList) {

			if (mc.getColspan() > 1 || mc.getRowspan() > 1) {

				cellIdList.add(mc.getId());
				String[] ids = mc.getId().split("_");
				int rowIndex = rowMap.get(ids[0]);
				int colIndex = colMap.get(ids[1]);
				for (int i = 0; i < mc.getRowspan(); i++) {
					for (int j = 0; j < mc.getColspan(); j++) {
						String row = sortRList.get(rowIndex + i).getAlias();
						String col = sortCList.get(colIndex + j).getAlias();
						MCell mcell = new MCell(mc);
						String id = row + "_" + col;
						mcell.setId(id);
						mcell.setColspan(1);
						mcell.setRowspan(1);
						mcell.getContent().setDisplayTexts(null);
						mcell.getContent().setTexts(null);
						tempList.add(mcell);

						mcellMap.put(mcell.getId(), mcell);
						// 存关系表
						MRowColCell mrcc = new MRowColCell();
						mrcc.setCellId(id);
						mrcc.setRow(row);
						mrcc.setCol(col);
						mrcc.setSheetId(sheetId);
						tempList.add(mrcc);
					}
				}

			} else {
				mc.getContent().setDisplayTexts(null);
				mc.getContent().setTexts(null);
				mcellMap.put(mc.getId(), mc);
				unspanCellIdList.add(mc.getId());
			}
		}
		// 删除合并单元格的MCell和关系表
		mrowColCellDao.delMRowColCellList1(excelId, sheetId, cellIdList);
		mcellDao.delMCell(excelId, sheetId, cellIdList);
		// 存入拆分之后的单元格及关系表
		baseDao.insertList(excelId, tempList);
		// 清除非合并单元格的显示内容
		mcellDao.updateContent("displayTexts", null, null, null,
				unspanCellIdList, excelId, sheetId);

		// 将数据赋值给目标区
		tempList.clear();// 清空，用于存贮新MCell和关系表
		List<OuterPasteData> datas = paste.getPasteData();
		for (OuterPasteData data : datas) {
			int rowIndex = oprRow + data.getRow();
			int colIndex = oprCol + data.getCol();
			String row = sortRList.get(rowIndex).getAlias();
			String col = sortCList.get(colIndex).getAlias();
			String id = row + "_" + col;
			MCell mc = mcellMap.get(id);
			if (null == mc) {
				mc = new MCell(id, sheetId);
				mc.getContent().setDisplayTexts(data.getContent());
				mc.getContent().setTexts(data.getContent());
				tempList.add(mc);
				MRowColCell mrcc = new MRowColCell();
				mrcc.setCellId(id);
				mrcc.setRow(row);
				mrcc.setCol(col);
				mrcc.setSheetId(sheetId);
				tempList.add(mrcc);
			} else {
				mc.getContent().setDisplayTexts(data.getContent());
				mc.getContent().setTexts(data.getContent());
				baseDao.update(excelId, mc);
			}
		}
		baseDao.insertList(excelId, tempList);
		msheetDao.updateStep(excelId, sheetId, step);

	}

	/**
	 * 是否可以粘贴
	 * 
	 * @param isAblePaste
	 * @param excelBook
	 * @return
	 */

	public boolean isAblePaste(OuterPaste outerPaste, String excelId) {
		String sheetId = excelId + 0;
		/*
		 * List<RowCol> sortRList = new ArrayList<RowCol>(); List<RowCol>
		 * sortCList = new ArrayList<RowCol>(); mrowColDao.getRowList(sortRList,
		 * excelId, sheetId); mrowColDao.getColList(sortCList, excelId,
		 * sheetId);
		 */
		int oprRow = outerPaste.getOprRow();
		int oprCol = outerPaste.getOprCol();
		Paste paste = pastDataAnalysis(outerPaste.getContent());
		int rowLen = paste.getRowLen() - 1;
		int colLen = paste.getColLen() - 1;
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		boolean canPaste = isPaste(excelId, sheetId, oprRow, oprCol, rowLen,
				colLen, msheet);
		return canPaste;
	}

	@Override
	public boolean isCopy(Copy copy, String excelId) {

		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getOprRow();
		int targetStartColIndex = copy.getTarget().getOprCol();
		String sheetId = excelId + 0;
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		boolean canPaste = isPaste(excelId, sheetId, targetStartRowIndex,
				targetStartColIndex, rowRange, colRange, msheet);
		return canPaste;
	}

	public boolean isCut(Copy copy, String excelId) {

		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getOprRow();
		int targetStartColIndex = copy.getTarget().getOprCol();
		String sheetId = excelId + 0;
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);

		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		List<String> rowId = new ArrayList<String>();
		for (int i = startRowIndex; i < endRowIndex + 1; i++) {
			rowId.add(sortRList.get(i).getAlias());
		}
		List<String> colId = new ArrayList<String>();
		for (int i = startColIndex; i < endColIndex + 1; i++) {
			colId.add(sortCList.get(i).getAlias());
		}
		List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, rowId, colId);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relation) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		if((null!=msheet.getPasswd())&&(msheet.getProtect())){
			for(MCell mc:mcellList){
				if((null == mc.getContent().getLocked())
						|| (mc.getContent().getLocked())){
					return false;
				}
			}
		}

		boolean canPaste = isCutTo(sortRList,sortCList,excelId, sheetId, targetStartRowIndex,
				targetStartColIndex, rowRange, colRange,msheet);
		return canPaste;
	}
	
	private boolean isCutTo(List<RowCol> sortRList,List<RowCol> sortCList,String excelId, String sheetId,
			int targetStartRowIndex, int targetStartColIndex, int rowRange,
			int colRange, MSheet msheet) {
		boolean canPaste = true;
		
		Map<String, Integer> rowMap = new HashMap<String, Integer>();
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		
		for (int i = 0; i < sortRList.size(); i++) {
			rowMap.put(sortRList.get(i).getAlias(), i);
		}
		
		for (int i = 0; i < sortCList.size(); i++) {
			colMap.put(sortCList.get(i).getAlias(), i);
		}

		List<String> rowList = new ArrayList<String>();
		List<String> colList = new ArrayList<String>();
		// 起始位置，超出行
		if (targetStartRowIndex > sortRList.size() - 1) {
			canPaste = false;
			return canPaste;
		}

		if (sortRList.size() < targetStartRowIndex + rowRange + 1) {
			for (int i = targetStartRowIndex; i < sortRList.size(); i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
		} else {
			for (int i = targetStartRowIndex; i < targetStartRowIndex + rowRange
					+ 1; i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
		}
		// 起始位置，超出列
		if (targetStartColIndex > sortCList.size() - 1) {
			canPaste = false;
			return canPaste;
		}

		if (sortCList.size() < targetStartColIndex + colRange + 1) {
			for (int i = targetStartColIndex; i < sortCList.size(); i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		} else {
			for (int i = targetStartColIndex; i < targetStartColIndex + colRange
					+ 1; i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		}

		List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relation) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);

		for (MCell mc : mcellList) {

			if (mc.getColspan() > 1 || mc.getRowspan() > 1) {
				String[] ids = mc.getId().split("_");
				int rowIndex = rowMap.get(ids[0]);
				int colIndex = colMap.get(ids[1]);
				if (rowIndex + mc.getRowspan() > targetStartRowIndex + rowRange
						+ 1) {
					canPaste = false;
					break;
				}
				if (colIndex + mc.getColspan() > targetStartColIndex + colRange
						+ 1) {
					canPaste = false;
					break;
				}
			}

			if ((null!=msheet.getPasswd())&&(msheet.getProtect())) {
				if ((null == mc.getContent().getLocked())
						|| (mc.getContent().getLocked())) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;

	}

	private boolean isPaste(String excelId, String sheetId,
			int targetStartRowIndex, int targetStartColIndex, int rowRange,
			int colRange, MSheet msheet) {
		boolean canPaste = true;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rowMap = new HashMap<String, Integer>();
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			rowMap.put(sortRList.get(i).getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortCList.size(); i++) {
			colMap.put(sortCList.get(i).getAlias(), i);
		}

		List<String> rowList = new ArrayList<String>();
		List<String> colList = new ArrayList<String>();
		// 起始位置，超出行
		if (targetStartRowIndex > sortRList.size() - 1) {
			canPaste = false;
			return canPaste;
		}

		if (sortRList.size() < targetStartRowIndex + rowRange + 1) {
			for (int i = targetStartRowIndex; i < sortRList.size(); i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
		} else {
			for (int i = targetStartRowIndex; i < targetStartRowIndex + rowRange
					+ 1; i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
		}
		// 起始位置，超出列
		if (targetStartColIndex > sortCList.size() - 1) {
			canPaste = false;
			return canPaste;
		}

		if (sortCList.size() < targetStartColIndex + colRange + 1) {
			for (int i = targetStartColIndex; i < sortCList.size(); i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		} else {
			for (int i = targetStartColIndex; i < targetStartColIndex + colRange
					+ 1; i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		}

		List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relation) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);

		for (MCell mc : mcellList) {

			if (mc.getColspan() > 1 || mc.getRowspan() > 1) {
				String[] ids = mc.getId().split("_");
				int rowIndex = rowMap.get(ids[0]);
				int colIndex = colMap.get(ids[1]);
				if (rowIndex + mc.getRowspan() > targetStartRowIndex + rowRange
						+ 1) {
					canPaste = false;
					break;
				}
				if (colIndex + mc.getColspan() > targetStartColIndex + colRange
						+ 1) {
					canPaste = false;
					break;
				}
			}

			if ((null!=msheet.getPasswd())&&(msheet.getProtect())) {
				if ((null == mc.getContent().getLocked())
						|| (mc.getContent().getLocked())) {
					canPaste = false;
					break;
				}
			}
		}
		return canPaste;

	}

	@Override
	public void cut(Copy copy, String excelId, Integer step) {
		String sheetId = excelId + 0;
		copyOrCut(copy, excelId, sheetId, step, 1);
	}

	@Override
	public void copy(Copy copy, String excelId, Integer step) {
		String sheetId = excelId + 0;
		copyOrCut(copy, excelId, sheetId, step, 0);
	}

	/**
	 * 用于处理剪切复制或复制
	 * 
	 * @param copy
	 * @param excelId
	 * @param sheetId
	 * @param step
	 * @param type
	 *            0 复制 1 剪切复制
	 */
	private void copyOrCut(Copy copy, String excelId, String sheetId,
			Integer step, int type) {
		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getOprRow();
		int targetStartColIndex = copy.getTarget().getOprCol();
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rowMap = new HashMap<String, Integer>();
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			rowMap.put(sortRList.get(i).getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortCList.size(); i++) {
			colMap.put(sortCList.get(i).getAlias(), i);
		}
		List<String> rowList = new ArrayList<String>();
		List<String> colList = new ArrayList<String>();
		for (int i = startRowIndex; i < endRowIndex + 1; i++) {
			rowList.add(sortRList.get(i).getAlias());
		}
		for (int i = startColIndex; i < endColIndex + 1; i++) {
			colList.add(sortCList.get(i).getAlias());
		}
		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
		}
		// 找到复制的数据
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		for (MCell mc : mcellList) {
			String[] ids = mc.getId().split("_");
			int row = rowMap.get(ids[0]) - startRowIndex;
			int col = colMap.get(ids[1]) - startColIndex;
			mc.setRow(row);
			mc.setCol(col);
		}

		if (type == 1) {
			// 清空复制区
			mrowColCellDao.delMRowColCellList(excelId, sheetId, rowList,
					colList);
			mcellDao.delMCell(excelId, sheetId, cellIdList);
		}

		// 清空，用于存贮目标区域的对象
		rowList.clear();
		colList.clear();
		relationList.clear();
		cellIdList.clear();

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		int maxRow = msheet.getMaxrow();
		int maxCol = msheet.getMaxcol();

		if (sortRList.size() < targetStartRowIndex + rowRange + 1) {
			for (int i = targetStartRowIndex; i < sortRList.size(); i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
			// 增加新的行
			for (int i = 0; i < targetStartRowIndex + rowRange + 1
					- sortRList.size(); i++) {
				maxRow = maxRow + 1;
				String row = maxRow + "";
				MRow mrow = new MRow(row, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(row);
				rowCol.setLength(19);
				if (i == 0) {
					rowCol.setPreAlias(
							sortRList.get(sortRList.size() - 1).getAlias());
				} else {
					rowCol.setPreAlias((maxRow - 1) + "");
				}
				sortRList.add(rowCol);
				// 存入简化的行
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "rList");
				// 存入行样式
				baseDao.insert(excelId, mrow);
			}
		} else {
			for (int i = targetStartRowIndex; i < targetStartRowIndex + rowRange
					+ 1; i++) {
				rowList.add(sortRList.get(i).getAlias());
			}

		}

		if (sortCList.size() < targetStartColIndex + colRange + 1) {

			for (int i = targetStartColIndex; i < sortCList.size(); i++) {
				colList.add(sortCList.get(i).getAlias());
			}

			// 增加新的列
			for (int i = 0; i < targetStartColIndex + colRange + 1
					- sortCList.size(); i++) {
				maxCol = maxCol + 1;
				String col = maxCol + "";
				MCol mcol = new MCol(col, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(col);
				rowCol.setLength(71);
				if (i == 0) {
					rowCol.setPreAlias(
							sortCList.get(sortCList.size() - 1).getAlias());
				} else {
					rowCol.setPreAlias((maxCol - 1) + "");
				}
				sortCList.add(rowCol);
				// 存入简化的行
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
				// 存入列样式
				baseDao.insert(excelId, mcol);
			}
		} else {
			for (int i = targetStartColIndex; i < targetStartColIndex + colRange
					+ 1; i++) {
				colList.add(sortCList.get(i).getAlias());
			}
		}
		msheet.setMaxcol(maxCol);
		msheet.setMaxrow(maxRow);
		baseDao.update(excelId, msheet);// 更新最大行，最大列
		relationList = mrowColCellDao.getMRowColCellList(excelId, sheetId,
				rowList, colList);
		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
		}
		// 删除目标区域的MCell对象
		mcellDao.delMCell(excelId, sheetId, cellIdList);
		// 删除目标区域关系对象
		mrowColCellDao.delMRowColCellList(excelId, sheetId, rowList, colList);
		// 在目标区写入复制数据
		List<Object> tempList = new ArrayList<Object>();
		for (MCell mc : mcellList) {
			int rowIndex = targetStartRowIndex + mc.getRow();
			int colIndex = targetStartColIndex + mc.getCol();
			String row = sortRList.get(rowIndex).getAlias();
			String col = sortCList.get(colIndex).getAlias();
			String id = row + "_" + col;
			mc.setId(id);
			mc.setRow(null);
			mc.setCol(null);
			tempList.add(mc);
			for (int i = 0; i < mc.getRowspan(); i++) {
				for (int j = 0; j < mc.getColspan(); j++) {
					MRowColCell mrcc = new MRowColCell();
					mrcc.setCellId(id);
					mrcc.setRow(sortRList.get(rowIndex + i).getAlias());
					mrcc.setCol(sortCList.get(colIndex + j).getAlias());
					mrcc.setSheetId(sheetId);

					tempList.add(mrcc);
				}
			}
		}
		baseDao.insertList(excelId, tempList);
		// 更新步骤
		msheetDao.updateStep(excelId, sheetId, step);

	}

	private Paste pastDataAnalysis(String data) {
		Paste paste = new Paste();
		String[] rowData = data.split("\n");
		paste.setRowLen(rowData.length);
		List<OuterPasteData> contentList = new ArrayList<OuterPasteData>();
		for (int i = 0; i < rowData.length; i++) {
			String[] colData = rowData[i].split("\t");
			for (int j = 0; j < colData.length; j++) {
				if (j == 0) {
					paste.setColLen(colData.length);
				}
				OuterPasteData outData = new OuterPasteData();
				outData.setRow(i);
				outData.setCol(j);
				outData.setContent(colData[j]);
				contentList.add(outData);
			}
		}
		paste.setPasteData(contentList);
		return paste;
	}

	@Override
	public void redo(String excelId) {
		HistoryCache cache = (HistoryCache) redis.get(excelId);
		if (null == cache) {
			return;
		}
		List<History> list = cache.getList();
		int index = cache.getIndex();
		flag = cache.getFlag();

		if ("redo".equals(flag)) {
			if (index == list.size()) {
				return;
			}
			index++;
			cache.setIndex(index);
		}

		History history = list.get(index - 1);
		Record record = history.getRecord();
		Before before = history.getBefore();
		if (record.getSure() == 1) {
			List<Param> params = record.getParam();
			for (Param param : params) {
				if ("MCellDaoImpl".equals(param.getTarget())) {
					Execute.exec(mcellDao, param.getName(),
							param.getArguments());
				} else if ("MColDaoImpl".equals(param.getTarget())) {
					Execute.exec(mcolDao, param.getName(),
							param.getArguments());
				} else if ("MRowColCellDaoImpl".equals(param.getTarget())) {
					Execute.exec(mrowColCellDao, param.getName(),
							param.getArguments());
				} else if ("MRowColDaoImpl".equals(param.getTarget())) {
					Execute.exec(mrowColDao, param.getName(),
							param.getArguments());
				} else if ("MRowDaoImpl".equals(param.getTarget())) {
					Execute.exec(mrowDao, param.getName(),
							param.getArguments());
				} else if ("BaseDao".equals(param.getTarget())) {
					Execute.exec(baseDao, param.getName(),
							param.getArguments());
				}
			}
			record.setSure(2);
			before.setSure(1);
		}
		cache.setFlag("redo");
		redis.set(excelId, cache);

	}

	@Override
	public void undo(String excelId) {
		HistoryCache cache = (HistoryCache) redis.get(excelId);
		if (null == cache) {
			return;
		}

		List<History> list = cache.getList();
		int index = cache.getIndex();
		flag = cache.getFlag();

		if ("undo".equals(flag)) {
			if (index == 1) {
				return;
			}
			index--;
			cache.setIndex(index);
		}

		History history = list.get(index - 1);

		String bookId = history.getExcelId();
		String sheetId = bookId + 0;

		Before before = history.getBefore();
		Record record = history.getRecord();

		if (before.getSure() == 1) {
			List<Object> delList = before.getDelList();
			for (Object o : delList) {
				baseDao.del(bookId, o);
			}

			List<MRowColList> rowList = before.getRowList();
			if (rowList.size() > 0) {
				MRowColList mrcl = rowList.get(0);
				mrowColDao.delRowColList(excelId, sheetId, mrcl.getId());
				baseDao.insert(bookId, rowList.get(0));
			}
			List<MRowColList> colList = before.getColList();
			if (colList.size() > 0) {
				MRowColList mrcl = colList.get(0);
				mrowColDao.delRowColList(excelId, sheetId, mrcl.getId());
				baseDao.insert(bookId, colList.get(0));
			}

			MColBefore mcol = before.getMcol();
			mcolDao.delMColList(excelId, sheetId, mcol.getIdList());

			baseDao.insertList(excelId, mcol.getMcolList());

			MRowBefore mrow = before.getMrow();
			mrowDao.delMRowList(excelId, sheetId, mrow.getIdList());
			baseDao.insertList(excelId, mrow.getMrowList());

			MCellBefore mcell = before.getMcell();
			mcellDao.delMCell(excelId, sheetId, mcell.getIdList());
			baseDao.insertList(excelId, mcell.getMcellList());

			MRowColCellBefore mrowColCell = before.getMrowColCell();
			mrowColCellDao.delMRowColCellList1(excelId, sheetId,
					mrowColCell.getIdList());
			baseDao.insertList(excelId, mrowColCell.getMrowColCellList());

			before.setSure(2);
			record.setSure(1);
		}
		cache.setFlag("undo");
		redis.set(excelId, cache);

	}

	@Override
	public void updateProtect(String excelId, MSheet msheet) {
		String sheetId = excelId + 0;
		if (msheet.getProtect()) {
			msheetDao.updateMSheetByObject(excelId, sheetId, msheet);
		} else {
			MSheet ms = msheetDao.getMSheet(excelId, sheetId);
			if (null == ms.getPasswd()
					|| msheet.getPasswd().equals(ms.getPasswd())) {
				msheetDao.updateMSheetByObject(excelId, sheetId, msheet);
			}
		}

	}

}
