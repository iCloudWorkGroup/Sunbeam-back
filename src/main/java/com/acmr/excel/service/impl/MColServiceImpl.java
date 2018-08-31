package com.acmr.excel.service.impl;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MColDao;
import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.distribute.Distribute;
import com.acmr.excel.distribute.Target;
import com.acmr.excel.model.ColOperate;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MColService;

@Service("mcolService")
public class MColServiceImpl implements MColService {

	@Resource
	private BaseDao baseDao;

	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;

	@Resource
	private MSheetDao msheetDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MRowColCellDao mrowColCellDao;

	@Override
	public void insertCol(ColOperate colOperate, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId, sheetId);
		mrowColDao.getRowList(sortRcList, excelId, sheetId);
		if (colOperate.getCol() > sortClList.size() - 1) {
			return;
		}
		Map<String, Integer> rowMap = new HashMap<String, Integer>();

		for (int i = 0; i < sortRcList.size(); i++) {
			RowCol rc = sortRcList.get(i);
			rowMap.put(rc.getAlias(), i);
		}
		
		List<Object> insertList = new ArrayList<Object>();// 存需要插入的对象
		
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		String alias = msheet.getMaxcol() + 1 + "";

		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		// 存一个列样式
		MCol mcol = new MCol();
		mcol.setSheetId(sheetId);
		mcol.setAlias(alias);
		
		if (colOperate.getCol() == 0) {
			rowCol.setPreAlias(null);
			rowCol.setLength(71);
			
			mcol.setHidden(false);
			mcol.setWidth(71);
		} else {
			String preAlias = sortClList.get(colOperate.getCol() - 1).getAlias();
			MCol preMCol = mcolDao.getMCol(excelId, sheetId, preAlias);
			
			rowCol.setPreAlias(
					sortClList.get(colOperate.getCol() - 1).getAlias());
			rowCol.setLength(preMCol.getWidth());
			
			mcol.setHidden(preMCol.getHidden());
			mcol.setWidth(preMCol.getWidth());
			setMColProperty(mcol,preMCol.getProps().getContent(),preMCol.getProps().getCustomProp());
			
			//获取前一列所有单元格,并将其content中属性赋值给新插入列的单元格
			List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId, sheetId, preAlias, "col");
			List<String> cellIdList = new ArrayList<String>();
			for(MRowColCell mrcc:relation){
				cellIdList.add(mrcc.getCellId());
			}
			List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId, cellIdList);
			for(MCell mc:mcellList){
				String[] ids = mc.getId().split("_");
				if(mc.getColspan()==1&&mc.getRowspan()==1){
					MCell mcell = new MCell(mc.getContent());
					mcell.setSheetId(sheetId);
					mcell.setId(ids[0]+"_"+alias);
					mcell.setRowspan(1);
					mcell.setColspan(1);
					
					mcell.getContent().setTexts(null);
					mcell.getContent().setDisplayTexts(null);
					mcell.getContent().setType(null);
					mcell.getContent().setExpress(null);
					insertList.add(mcell);
					
					MRowColCell mrcc = new MRowColCell();
					mrcc.setCellId(ids[0]+"_"+alias);
					mrcc.setRow(ids[0]);
					mrcc.setCol(alias);
					mrcc.setSheetId(sheetId);
					insertList.add(mrcc);// 关系表
				}else{
					int span = Integer.parseInt(preAlias)-Integer.parseInt(ids[1])+1;
					if(span==mc.getColspan()){
						String row = ids[0];
						Integer index = rowMap.get(row);
						for (int i = 0; i < mc.getRowspan(); i++) {
							MRowColCell mrcc = new MRowColCell();
							row = sortRcList.get(index).getAlias();
							index++;
							mrcc.setCellId(row+"_"+alias);
							mrcc.setRow(row);
							mrcc.setCol(alias);
							mrcc.setSheetId(sheetId);
							insertList.add(mrcc);
							
							MCell mcell = new MCell(mc.getContent());
							mcell.setSheetId(sheetId);
							mcell.setId(row+"_"+alias);
							mcell.setRowspan(1);
							mcell.setColspan(1);
							
							mcell.getContent().setTexts(null);
							mcell.getContent().setDisplayTexts(null);
							mcell.getContent().setType(null);
							mcell.getContent().setExpress(null);
							insertList.add(mcell);
						}
					}
				}
				
			}
		}

		mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
		// 修改选中行的前列别名
		String nextAlias = sortClList.get(colOperate.getCol()).getAlias();
		mrowColDao.updateRowCol(excelId, sheetId, "cList", nextAlias, alias);

		baseDao.insert(excelId, mcol);

		// 当选中行是当前可视行时
		if ((null != msheet.getFreeze())
				&& (msheet.getViewColAlias().equals(nextAlias))
				&& (msheet.getFreeze())) {
			msheetDao.updateMSheetProperty(excelId, sheetId, "viewColAlias",
					alias);
		}

		msheetDao.updateMaxCol(msheet.getMaxcol() + 1, excelId, sheetId);

		String col = sortClList.get(colOperate.getCol()).getAlias();
		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, col, "col");
		List<String> cellIdList = new ArrayList<>();

		for (MRowColCell rcc : relationList) {
			cellIdList.add(rcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);

		

		for (MCell mc : mcellList) {
			String[] ids = mc.getId().split("_");
			if ((mc.getColspan() > 1) && (!ids[1].equals(col))) {
				mc.setColspan(mc.getColspan() + 1);
				baseDao.update(excelId, mc);
				String row = ids[0];
				Integer index = rowMap.get(row);
				for (int i = 0; i < mc.getRowspan(); i++) {
					MRowColCell mrcc = new MRowColCell();
					row = sortRcList.get(index).getAlias();
					index++;
					mrcc.setCellId(mc.getId());
					mrcc.setRow(row);
					mrcc.setCol(alias);
					mrcc.setSheetId(sheetId);
					insertList.add(mrcc);
				}
			}

		}
		if (insertList.size() > 0) {
			baseDao.insertList(excelId, insertList);
		}
		msheetDao.updateStep(excelId, sheetId, step);
		/*
		 * Map<String,MCell> cellMap = new HashMap<String,MCell>(); for(MCell
		 * mec:mcellList){ cellMap.put(mec.getId(), mec); }
		 * 
		 * List<MCell> sortMCellList = new ArrayList<MCell>();
		 * 
		 * for(RowCol rc: sortRcList){ String cellId =
		 * relationMap.get(rc.getAlias()+"_"+col);
		 * sortMCellList.add(cellMap.get(cellId)); }
		 * 
		 * List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		 * for(int i=0;i<sortRcList.size();i++){ MCell mec =
		 * sortMCellList.get(i); RowCol rc = sortRcList.get(i); if(null == mec){
		 * continue; } if(null!=mec&&mec.getColspan()==1){
		 * //new新增一行的MExcelCell对象 MExcelCell me = new MExcelCell();
		 * me.setId(rc.getAlias()+"_"+alias); me.setRowId(rc.getAlias());
		 * me.setColId(alias); me.setRowspan(1); me.setColspan(1); ExcelCell ec
		 * = new ExcelCell(); me.setExcelCell(ec); tempList.add(me); RowColCell
		 * rcc = new RowColCell(); rcc.setCellId(rc.getAlias()+"_"+alias);
		 * rcc.setCol(Integer.parseInt(alias));
		 * rcc.setRow(Integer.parseInt(rc.getAlias())); tempList.add(rcc); }else
		 * if(mec.getColspan()>1){ if(mec.getColId().equals(colId) ){
		 * //new新增一行的MExcelCell对象,是以选定列作为开始列的合并单元格 MExcelCell me = new
		 * MExcelCell(); me.setId(rc.getAlias()+"_"+alias);
		 * me.setRowId(rc.getAlias()); me.setColId(alias); me.setRowspan(1);
		 * me.setColspan(1); ExcelCell ec = new ExcelCell();
		 * me.setExcelCell(ec); tempList.add(me); //关系对象 RowColCell rcc = new
		 * RowColCell(); rcc.setCellId(rc.getAlias()+"_"+alias);
		 * rcc.setCol(Integer.parseInt(alias));
		 * rcc.setRow(Integer.parseInt(rc.getAlias())); tempList.add(rcc);
		 * }else{ if(mec.getRowId().equals(rc.getAlias())){
		 * mec.setColspan(mec.getColspan()+1); tempList.add(mec); //关系对象
		 * RowColCell rcc = new RowColCell(); rcc.setCellId(mec.getId());
		 * rcc.setCol(Integer.parseInt(alias));
		 * rcc.setRow(Integer.parseInt(rc.getAlias())); tempList.add(rcc);
		 * }else{ //关系对象 RowColCell rcc = new RowColCell();
		 * rcc.setCellId(mec.getId()); rcc.setCol(Integer.parseInt(alias));
		 * rcc.setRow(Integer.parseInt(rc.getAlias())); tempList.add(rcc); } } }
		 * }
		 */
	}

	public void insertColSelf(Map<String, Object> map) {
		String excelId = (String) map.get("excelId");
		String sheetId = (String) map.get("sheetId");
		ColOperate colOperate = (ColOperate) map.get("colOperate");
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId, sheetId);
		mrowColDao.getRowList(sortRcList, excelId, sheetId);
		if (colOperate.getCol() > sortClList.size() - 1) {
			return;
		}
		Map<String, Integer> rowMap = new HashMap<String, Integer>();

		for (int i = 0; i < sortRcList.size(); i++) {
			RowCol rc = sortRcList.get(i);
			rowMap.put(rc.getAlias(), i);
		}

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		String alias = msheet.getMaxcol() + 1 + "";

		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		rowCol.setLength(71);
		if (colOperate.getCol() == 0) {
			rowCol.setPreAlias(null);
		} else {
			rowCol.setPreAlias(
					sortClList.get(colOperate.getCol() - 1).getAlias());
		}

		mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
		// 修改选中行的前列别名
		String col = sortClList.get(colOperate.getCol()).getAlias();
		mrowColDao.updateRowCol(excelId, sheetId, "cList", col, alias);
		// 存一个列样式
		MCol mcol = new MCol();
		mcol.setSheetId(sheetId);
		mcol.setAlias(alias);
		mcol.setHidden(false);
		mcol.setWidth(71);

		baseDao.insert(excelId, mcol);

		map.put("col", col);
		map.put("msheet", msheet);
		map.put("alias", alias);
		map.put("sortRcList", sortRcList);
		map.put("rowMap", rowMap);

	}

	public void insertColEffectMSheet(Map<String, Object> map) {
		MSheet msheet = (MSheet) map.get("msheet");
		String col = (String) map.get("col");
		String alias = (String) map.get("alias");
		String excelId = (String) map.get("excelId");
		String sheetId = (String) map.get("sheetId");
		// 当选中行是当前可视行时
		if ((null != msheet.getFreeze())
				&& (msheet.getViewColAlias().equals(col))
				&& (msheet.getFreeze())) {
			msheetDao.updateMSheetProperty(excelId, sheetId, "viewColAlias",
					alias);
		}

		msheetDao.updateMaxCol(msheet.getMaxcol() + 1, excelId, sheetId);

	}

	public void insertColEffectMCell(Map<String, Object> map) {
		// String col = sortClList.get(colOperate.getCol()).getAlias();
		String excelId = (String) map.get("excelId");
		String sheetId = (String) map.get("sheetId");
		@SuppressWarnings("unchecked")
		List<RowCol> sortRcList = (List<RowCol>) map.get("sortRcList");
		@SuppressWarnings("unchecked")
		Map<String, Integer> rowMap = (Map<String, Integer>) map.get("rowMap");
		String alias = (String) map.get("alias");
		Integer step = (Integer) map.get("step");
		String col = (String) map.get("col");

		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, col, "col");
		List<String> cellIdList = new ArrayList<>();

		for (MRowColCell rcc : relationList) {
			cellIdList.add(rcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);

		List<Object> insertList = new ArrayList<Object>();// 存需要插入的对象

		for (MCell mc : mcellList) {
			String[] ids = mc.getId().split("_");
			if ((mc.getColspan() > 1) && (!ids[1].equals(col))) {
				mc.setColspan(mc.getColspan() + 1);
				baseDao.update(excelId, mc);
				String row = ids[0];
				Integer index = rowMap.get(row);
				for (int i = 0; i < mc.getRowspan(); i++) {
					MRowColCell mrcc = new MRowColCell();
					row = sortRcList.get(index).getAlias();
					index++;
					mrcc.setCellId(mc.getId());
					mrcc.setRow(row);
					mrcc.setCol(alias);
					mrcc.setSheetId(sheetId);
					insertList.add(mrcc);
				}
			}

		}
		if (insertList.size() > 0) {
			baseDao.insertList(excelId, insertList);
		}
		msheetDao.updateStep(excelId, sheetId, step);

	}

	@Override
	public void addCol(int num, String excelId, Integer step) {
		String sheetId = excelId + 0;
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		int maxCol = msheet.getMaxcol();
		List<RowCol> sortCList = new ArrayList<>();
		mrowColDao.getColList(sortCList, excelId, sheetId);
		// 增加新的列
		for (int i = 0; i < num; i++) {
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
			// 存入简化的列
			mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
			// 存入列样式
			baseDao.insert(excelId, mcol);

		}
		msheet.setMaxcol(maxCol);
		msheet.setStep(step);
		baseDao.update(excelId, msheet);// 更新最大列及步骤
	}

	@Override
	public void delCol(ColOperate colOperate, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId, sheetId);
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		if (colOperate.getCol() > sortCList.size() - 1) {
			return;
		}
		// 列别名
		String alias = sortCList.get(colOperate.getCol()).getAlias();
		// 删除列
		mrowColDao.delRowCol(excelId, sheetId, alias, "cList");
		// 修改这列后面一列的前列别名
		int index = colOperate.getCol();
		String frontAlias;
		if (colOperate.getCol() == 0) {
			frontAlias = null;
		} else {
			frontAlias = sortCList.get(index - 1).getAlias();// 前一列别名
		}

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		String backAlias = null;
		// 如果删除的是最后一列，它没有后一行
		if (index != sortCList.size() - 1) {
			backAlias = sortCList.get(index + 1).getAlias();// 后一列别名
			mrowColDao.updateRowCol(excelId, sheetId, "cList", backAlias,
					frontAlias);
			// 当删除行是可视行与冻结行时
			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getViewColAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "viewColAlias",
						backAlias);

			}

			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getColAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "colAlias",
						backAlias);

			}
		} else {
			// 当删除行时冻结行时
			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getViewRowAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "viewColAlias",
						frontAlias);

			}

			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getColAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "colAlias",
						frontAlias);

			}

		}
		// 删除列样式记录
		mcolDao.delMCol(excelId, sheetId, alias);

		// 查找关系表
		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, alias, "col");
		List<String> cellIdList = new ArrayList<String>();
		// Map<String,String> relationMap = new HashMap<String,String>();
		for (MRowColCell mrcc : relationList) {

			// relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> cellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		// 删除关系表
		mrowColCellDao.delMRowColCell(excelId, sheetId, "col", alias);
		cellIdList.clear();// 存需要删除的MExcelCell的Id
		List<Object> tempList = new ArrayList<Object>();// 存需要修改或增加的MCell对象
		for (MCell mc : cellList) {
			if (mc.getColspan() == 1) {
				cellIdList.add(mc.getId());
			} else {
				String[] ids = mc.getId().split("_");
				if (ids[1].equals(alias)) {
					// 删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					String id = ids[0] + "_" + backAlias;
					// 修改合并单元格其他关系表的cellId
					mrowColCellDao.updateMRowColCell(excelId, sheetId,
							mec.getId(), id);
					mec.setId(id);
					mec.setColspan(mc.getColspan() - 1);

				} else {
					// 删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					mec.setColspan(mc.getColspan() - 1);
					tempList.add(mec);// 插入新MCell
				}
			}
		}

		mcellDao.delMCell(excelId, sheetId, cellIdList);
		if (tempList.size() > 0) {
			baseDao.insertList(excelId, tempList);

		}

		msheetDao.updateStep(excelId, sheetId, step);

		/*
		 * List<MExcelCell> sortMcellList = new ArrayList<MExcelCell>();
		 * Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
		 * for(MExcelCell ec:cellList){ cellMap.put(ec.getId(), ec); }
		 * for(RowCol rc:sortRList){ String cellId =
		 * relationMap.get(rc.getAlias()+"_"+alias);
		 * sortMcellList.add(cellMap.get(cellId)); }
		 * 
		 * 
		 * for(int i=0;i<sortRList.size();i++){ String row =
		 * sortRList.get(i).getAlias(); MExcelCell ec = sortMcellList.get(i);
		 * if(null == ec){ continue; } if(ec.getColspan() == 1){
		 * cellIdList.add(ec.getId()); }else{ if(ec.getColId().equals(alias)){
		 * if(ec.getRowId().equals(row)){ //删除老的MExcelCell
		 * cellIdList.add(ec.getId()); MExcelCell mec = ec;
		 * mec.setColId(backAlias); String id = row+"_"+backAlias;
		 * //修改合并单元格其他关系表的cellId mcellDao.updateRowColCell(excelId, mec.getId(),
		 * id); mec.setId(id); mec.setColspan(ec.getColspan()-1);
		 * 
		 * ExcelCell cell = ec.getExcelCell();
		 * cell.setColspan(cell.getColspan()-1); mec.setExcelCell(cell);
		 * cellList.add(mec);//插入新MExcelCell
		 * 
		 * } }else{ if(ec.getRowId().equals(row)){ MExcelCell mec = ec;
		 * mec.setColspan(mec.getColspan()-1); cellList.add(mec); } } }
		 * 
		 * }
		 */

	}

	@Override
	public void hideCol(ColOperate colOperate, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId, sheetId);
		String alias = sortCList.get(colOperate.getCol()).getAlias();

		mcolDao.updateColHiddenStatus(excelId, sheetId, alias, true);

		mrowColDao.updateRowColLength(excelId, sheetId, "cList", alias, 0);
		// mcellDao.updateHidden("colId", alias, true, excelId);
		msheetDao.updateStep(excelId, sheetId, step);
	}

	@Override
	public void showCol(ColOperate colOperate, String excelId, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId, sheetId);
		String alias = sortCList.get(colOperate.getCol()).getAlias();
		mcolDao.updateColHiddenStatus(excelId, sheetId, alias, false);

		MCol mcol = mcolDao.getMCol(excelId, sheetId, alias);
		mrowColDao.updateRowColLength(excelId, sheetId, "cList", alias,
				mcol.getWidth());
		// mcellDao.updateHidden("colId", alias, false, excelId);
		msheetDao.updateStep(excelId, sheetId, step);
	}

	@Override
	public void updateColWidth(ColWidth colWidth, String excelId,
			Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId, sheetId);
		String alias = sortCList.get(colWidth.getCol()).getAlias();
		Integer width = colWidth.getOffset();

		mcolDao.updateColWidth(excelId, sheetId, alias, width);
		mrowColDao.updateRowColLength(excelId, sheetId, "cList", alias, width);

		msheetDao.updateStep(excelId, sheetId, step);

	}
	
	private void setMColProperty(MCol mcol, Content content,
			CustomProp customProp) {
		try {
			Class<? extends Content> c = content.getClass();
			Content mc = mcol.getProps().getContent();
			Field[] cFields = c.getDeclaredFields();
			for (Field f : cFields) {
				f.setAccessible(true);
				Object value = f.get(content);
				if (null != value) {
					Field mf = mc.getClass().getDeclaredField(f.getName());
					mf.setAccessible(true);
					mf.set(mc, value);
				}
			}

			Class<? extends CustomProp> p = customProp.getClass();
			CustomProp mp = mcol.getProps().getCustomProp();
			Field[] pFields = p.getDeclaredFields();
			for (Field f : pFields) {
				f.setAccessible(true);
				Object value = f.get(customProp);
				if (null != value) {
					Field pf = mp.getClass().getDeclaredField(f.getName());
					pf.setAccessible(true);
					pf.set(mp, value);
				}
			}

		} catch (Exception e) {

			e.printStackTrace();
		}
	}

	@Override
	public void insertColDis(ColOperate colOperate, String excelId,
			Integer step) {
		// TODO Auto-generated method stub
		
	}

}
