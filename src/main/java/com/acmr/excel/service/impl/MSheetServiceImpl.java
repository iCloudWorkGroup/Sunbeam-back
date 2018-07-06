package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OuterPaste;
import com.acmr.excel.model.OuterPasteData;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MSheetService;

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

	@Override
	public void frozen(Frozen frozen, String excelId, Integer step) {
		String sheetId = excelId + 0;
		String oprRow = frozen.getOprRow();
		String oprCol = frozen.getOprCol();
		String viewRow = frozen.getViewRow();
		String viewCol = frozen.getViewCol();
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
		List<MRowColCell> relation = mcellDao.getMRowColCellList(excelId,
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
		mcellDao.delMRowColCellList(excelId, sheetId, cellIdList);
		mcellDao.delMCell(excelId, sheetId, cellIdList);
		// 存入拆分之后的单元格及关系表
		baseDao.insert(excelId, tempList);
		// 清除非合并单元格的显示内容
		mcellDao.updateContent("displayTexts", null,null,null, unspanCellIdList, excelId,
				sheetId);

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
		baseDao.insert(excelId, tempList);
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
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		int oprRow = outerPaste.getOprRow();
		int oprCol = outerPaste.getOprCol();
		Paste paste = pastDataAnalysis(outerPaste.getContent());
		int rowLen = paste.getRowLen() - 1;
		int colLen = paste.getColLen() - 1;
		boolean canPaste = isPaste(excelId, sheetId, oprRow, oprCol, rowLen,
				colLen);
		return canPaste;
	}

	@Override
	public boolean isCutCopy(Copy copy, String excelId) {

		int startRowIndex = copy.getOrignal().getStartRow();
		int endRowIndex = copy.getOrignal().getEndRow();
		int startColIndex = copy.getOrignal().getStartCol();
		int endColIndex = copy.getOrignal().getEndCol();
		int rowRange = endRowIndex - startRowIndex;
		int colRange = endColIndex - startColIndex;
		int targetStartRowIndex = copy.getTarget().getOprRow();
		int targetStartColIndex = copy.getTarget().getOprCol();
		String sheetId = excelId + 0;
		boolean canPaste = isPaste(excelId, sheetId, targetStartRowIndex,
				targetStartColIndex, rowRange, colRange);
		return canPaste;
	}

	private boolean isPaste(String excelId, String sheetId,
			int targetStartRowIndex, int targetStartColIndex, int rowRange,
			int colRange) {
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

		List<MRowColCell> relation = mcellDao.getMRowColCellList(excelId,
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

			if ((null != mc.getContent().getLocked())
					&& (mc.getContent().getLocked())) {
				canPaste = false;
				break;
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
		List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId,
				sheetId, rowList, colList);
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
			mcellDao.delMRowColCellList(excelId, sheetId, rowList, colList);
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
		relationList = mcellDao.getMRowColCellList(excelId, sheetId, rowList,
				colList);
		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
		}
		// 删除目标区域的MCell对象
		mcellDao.delMCell(excelId, sheetId, cellIdList);
		// 删除目标区域关系对象
		mcellDao.delMRowColCellList(excelId, sheetId, rowList, colList);
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
		baseDao.insert(excelId, tempList);
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
	public void delRow(String excelId, String sheetId) {
		// TODO Auto-generated method stub
		
	}

}
