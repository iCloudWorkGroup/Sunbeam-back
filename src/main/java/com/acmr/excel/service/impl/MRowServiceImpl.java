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
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowOperate;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MRowService;

@Service("mrowService")
public class MRowServiceImpl implements MRowService {

	@Resource
	private BaseDao baseDao;

	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;

	@Resource
	private MSheetDao msheetDao;
	@Resource
	private MRowDao mrowDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MRowColCellDao mrowColCellDao;

	public void insertRow(String excelId,RowOperate rowOperate,  Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId, sheetId);
		mrowColDao.getRowList(sortRcList, excelId, sheetId);
		if (rowOperate.getRow() > sortRcList.size() - 1) {
			return;
		}
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		for (int i = 0; i < sortClList.size(); i++) {
			RowCol rc = sortClList.get(i);
			colMap.put(rc.getAlias(), i);
		}
        
		
		
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		String alias = msheet.getMaxrow() + 1 + "";
		String viewRowAlias = msheet.getViewRowAlias();
		
		List<Object> tempList = new ArrayList<Object>();// 存需要插入的对象

		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		// 存一个行样式
		MRow mr = new MRow();
		mr.setAlias(alias);
		mr.setSheetId(sheetId);
		if (rowOperate.getRow() == 0) {
			rowCol.setPreAlias(null);
			rowCol.setLength(19);
			
			mr.setHeight(19);
			mr.setHidden(false);
		} else {
			String preAlias = sortRcList.get(rowOperate.getRow() - 1).getAlias();
			MRow preMRow = mrowDao.getMRow(excelId, sheetId, preAlias);
			
			rowCol.setPreAlias(
					sortRcList.get(rowOperate.getRow() - 1).getAlias());
			rowCol.setLength(preMRow.getHeight());
			
			mr.setHeight(preMRow.getHeight());
			mr.setHidden(preMRow.getHidden());
			setMRowProperty(mr,preMRow.getProps().getContent(),preMRow.getProps().getCustomProp());
			
			//获取前一行所有单元格
			List<MRowColCell> relation = mrowColCellDao.getMRowColCellList(excelId, sheetId, preAlias, "row");
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
					mcell.setId(alias+"_"+ids[1]);
					mcell.setRowspan(1);
					mcell.setColspan(1);
					
					mcell.getContent().setTexts(null);
					mcell.getContent().setDisplayTexts(null);
					mcell.getContent().setType(null);
					mcell.getContent().setExpress(null);
					tempList.add(mcell);
					
					MRowColCell mrcc = new MRowColCell();
					mrcc.setCellId(alias+"_"+ids[1]);
					mrcc.setRow(alias);
					mrcc.setCol(ids[1]);
					mrcc.setSheetId(sheetId);
					tempList.add(mrcc);// 关系表
				}else{
					int span = Integer.parseInt(preAlias)-Integer.parseInt(ids[0])+1;
					if(span==mc.getRowspan()){
						String col = ids[1];
						Integer index = colMap.get(col);
						for (int i = 0; i < mc.getColspan(); i++) {
							MRowColCell mrcc = new MRowColCell();
							col = sortClList.get(index).getAlias();
							index++;
							mrcc.setCellId(alias+"_"+col);
							mrcc.setRow(alias);
							mrcc.setCol(col);
							mrcc.setSheetId(sheetId);
							tempList.add(mrcc);
							
							MCell mcell = new MCell(mc.getContent());
							mcell.setSheetId(sheetId);
							mcell.setId(alias+"_"+col);
							mcell.setRowspan(1);
							mcell.setColspan(1);
							
							mcell.getContent().setTexts(null);
							mcell.getContent().setDisplayTexts(null);
							mcell.getContent().setType(null);
							mcell.getContent().setExpress(null);
							tempList.add(mcell);
						}
					}
				}
				
			}
			
		}

		mrowColDao.insertRowCol(excelId, sheetId, rowCol, "rList");
		// 修改选中行的前行别名
		String nextAlias = sortRcList.get(rowOperate.getRow()).getAlias();
		mrowColDao.updateRowCol(excelId, sheetId, "rList", nextAlias, alias);
		baseDao.insert(excelId, mr);//存样行样式
		// 当选中行是当前可视行时
		if ((null != msheet.getFreeze()) && (viewRowAlias.equals(nextAlias))
				&& (msheet.getFreeze())) {
			msheetDao.updateMSheetProperty(excelId, sheetId, "viewRowAlias",
					alias);
		}

		msheetDao.updateMaxRow(msheet.getMaxrow() + 1, excelId, sheetId);

		String row = sortRcList.get(rowOperate.getRow()).getAlias();
		List<MRowColCell> relationList = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, row, "row");
		List<String> cellIdList = new ArrayList<>();

		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		
		for (MCell mc : mcellList) {
			String[] ids = mc.getId().split("_");
			if ((mc.getRowspan() > 1) && (!ids[0].equals(row))) {
				mc.setRowspan(mc.getRowspan() + 1);
				baseDao.update(excelId, mc);
				String col = ids[1];
				Integer index = colMap.get(col);
				for (int i = 0; i < mc.getColspan(); i++) {
					MRowColCell mrcc = new MRowColCell();

					col = sortClList.get(index).getAlias();
					index++;
					mrcc.setCellId(mc.getId());
					mrcc.setRow(alias);
					mrcc.setCol(col);
					mrcc.setSheetId(sheetId);
					tempList.add(mrcc);
				}
			}

		}

		if (tempList.size() > 0) {

			baseDao.insertList(excelId, tempList);
		}

		msheetDao.updateStep(excelId, sheetId, step);

		/*
		 * Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
		 * for(MExcelCell mec:mcellList){ cellMap.put(mec.getId(), mec); }
		 * String rowId = sortRcList.get(rowOperate.getRow()).getAlias();//行别名
		 * List<MExcelCell> sortMCellList = new ArrayList<MExcelCell>();
		 * 
		 * for(RowCol rc: sortClList){ String cellId =
		 * relationMap.get(rowId+"_"+rc.getAlias());
		 * sortMCellList.add(cellMap.get(cellId)); }
		 * 
		 * List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		 * for(int i=0;i<sortClList.size();i++){ MExcelCell mec =
		 * sortMCellList.get(i); if(null==mec){ continue; } RowCol rc =
		 * sortClList.get(i); if(mec.getRowspan()==1){ //new新增一行的MExcelCell对象
		 * MExcelCell me = new MExcelCell(); me.setId(alias+"_"+rc.getAlias());
		 * me.setRowId(alias); me.setColId(rc.getAlias()); me.setRowspan(1);
		 * me.setColspan(1); ExcelCell ec = new ExcelCell();
		 * me.setExcelCell(ec); tempList.add(me); RowColCell rcc = new
		 * RowColCell(); rcc.setCellId(alias+"_"+rc.getAlias());
		 * rcc.setCol(Integer.parseInt(rc.getAlias()));
		 * rcc.setRow(Integer.parseInt(alias)); tempList.add(rcc); }else
		 * if(mec.getRowspan()>1){ if(mec.getRowId().equals(rowId) ){
		 * //new新增一行的MExcelCell对象,是以选定行作为开始行的合并单元格 MExcelCell me = new
		 * MExcelCell(); me.setId(alias+"_"+rc.getAlias()); me.setRowId(alias);
		 * me.setColId(rc.getAlias()); me.setRowspan(1); me.setColspan(1);
		 * ExcelCell ec = new ExcelCell(); me.setExcelCell(ec);
		 * tempList.add(me); //关系对象 RowColCell rcc = new RowColCell();
		 * rcc.setCellId(alias+"_"+rc.getAlias());
		 * rcc.setCol(Integer.parseInt(rc.getAlias()));
		 * rcc.setRow(Integer.parseInt(alias)); tempList.add(rcc); }else{
		 * if(mec.getColId().equals(rc.getAlias())){
		 * mec.setRowspan(mec.getRowspan()+1); tempList.add(mec); //关系对象
		 * RowColCell rcc = new RowColCell(); rcc.setCellId(mec.getId());
		 * rcc.setCol(Integer.parseInt(rc.getAlias()));
		 * rcc.setRow(Integer.parseInt(alias)); tempList.add(rcc); }else{ //关系对象
		 * RowColCell rcc = new RowColCell(); rcc.setCellId(mec.getId());
		 * rcc.setCol(Integer.parseInt(rc.getAlias()));
		 * rcc.setRow(Integer.parseInt(alias)); tempList.add(rcc); } } } }
		 */

	}

	@Override
	public void addRow(String excelId,int num,  Integer step) {
		String sheetId = excelId + 0;
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		int maxRow = msheet.getMaxrow();
		List<RowCol> sortRList = new ArrayList<>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		// 增加新的行
		for (int i = 0; i < num; i++) {
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
			// 存入简化的行
			mrowColDao.insertRowCol(excelId, sheetId, rowCol, "rList");
			// 存入行样式
			baseDao.insert(excelId, mrow);
		}
		msheet.setMaxrow(maxRow);
		msheet.setStep(step);
		baseDao.update(excelId, msheet);// 更新最大列及步骤
		
	}

	@Override
	public void delRow(String excelId,RowOperate rowOperate,  Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		if (rowOperate.getRow() > sortRList.size() - 1) {
			return;
		}
		// 行别名
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		// 删除行
		mrowColDao.delRowCol(excelId, sheetId, alias, "rList");
		// 修改这行后面一行的前一行别名
		int index = rowOperate.getRow();
		String frontAlias;
		if (rowOperate.getRow() == 0) {
			frontAlias = null;// 前一行别名
		} else {
			frontAlias = sortRList.get(index - 1).getAlias();// 前一行别名
		}

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		String backAlias = null;
		// 当删除的是最后一行，不用修改它后面一行的前行别名
		if (index != sortRList.size() - 1) {
			backAlias = sortRList.get(index + 1).getAlias();// 下一行别名
			mrowColDao.updateRowCol(excelId, sheetId, "rList", backAlias,
					frontAlias);

			// 当删除行是可视行与冻结行时
			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getViewRowAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "viewRowAlias",
						backAlias);

			}

			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getRowAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "rowAlias",
						backAlias);

			}
		} else {
			// 当删除行时冻结行时
			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getViewRowAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "viewRowAlias",
						frontAlias);

			}

			if ((null != msheet.getFreeze())
					&& (alias.equals(msheet.getRowAlias())
							&& (msheet.getFreeze()))) {
				msheetDao.updateMSheetProperty(excelId, sheetId, "rowAlias",
						frontAlias);

			}

		}
		// 删除行样式记录
		mrowDao.delMRow(excelId, sheetId, alias);

		List<MRowColCell> relationList = mrowColCellDao.getMRowColCellList(excelId,
				sheetId, alias, "row");
		List<String> cellIdList = new ArrayList<String>();

		for (MRowColCell mrcc : relationList) {

			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> cellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		// 删除关系表
		mrowColCellDao.delMRowColCell(excelId, sheetId, "row", alias);
		cellIdList.clear();// 存需要删除的MExcelCell的Id
		List<Object> tempList = new ArrayList<Object>();// 存需要修改或增加的MCell对象
		for (MCell mc : cellList) {
			if (mc.getRowspan() == 1) {
				// 删除老的MExcelCell
				cellIdList.add(mc.getId());
			} else {
				String[] ids = mc.getId().split("_");
				if (ids[0].equals(alias)) {
					// 删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					String id = backAlias + "_" + ids[1];
					// 修改合并单元格其他关系表的cellId
					mrowColCellDao.updateMRowColCell(excelId, sheetId, mec.getId(),
							id);
					mec.setId(id);
					mec.setRowspan(mc.getRowspan() - 1);

					tempList.add(mec);// 插入新MCell
				} else {
					// 删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					mec.setRowspan(mc.getRowspan() - 1);
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
		 * for(RowCol rc:sortCList){ String cellId =
		 * relationMap.get(alias+"_"+rc.getAlias());
		 * sortMcellList.add(cellMap.get(cellId)); } //删除关系表
		 * mcellDao.delRowColCell(excelId,"row",Integer.parseInt(alias));
		 * cellIdList.clear();//存需要删除的MExcelCell的Id
		 * cellList.clear();//存需要修改或增加的MExcelCell对象 for(int
		 * i=0;i<sortCList.size();i++){ String col =
		 * sortCList.get(i).getAlias(); MExcelCell ec = sortMcellList.get(i);
		 * if(null == ec){ continue; } if(ec.getRowspan() == 1){
		 * cellIdList.add(ec.getId()); }else{ if(ec.getRowId().equals(alias)){
		 * if(ec.getColId().equals(col)){ //删除老的MExcelCell
		 * cellIdList.add(ec.getId()); MExcelCell mec = ec;
		 * mec.setRowId(backAlias); String id = backAlias+"_"+col;
		 * //修改合并单元格其他关系表的cellId mcellDao.updateRowColCell(excelId, mec.getId(),
		 * id); mec.setId(backAlias+"_"+col); mec.setRowspan(ec.getRowspan()-1);
		 * 
		 * ExcelCell cell = ec.getExcelCell();
		 * cell.setRowspan(cell.getRowspan()-1); mec.setExcelCell(cell);
		 * cellList.add(mec);//插入新MExcelCell
		 * 
		 * } }else{ if(ec.getColId().equals(col)){ MExcelCell mec = ec;
		 * mec.setRowspan(mec.getRowspan()-1); cellList.add(mec); } } }
		 * 
		 * }
		 */

	}

	@Override
	public void hideRow( String excelId,RowOperate rowOperate, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		mrowDao.updateRowHidden(excelId, sheetId, alias, true);
		mrowColDao.updateRowColLength(excelId, sheetId, "rList", alias, 0);
		// mcellDao.updateHidden("rowId", alias, true, excelId);
		msheetDao.updateStep(excelId, sheetId, step);

	}

	@Override
	public void showRow( String excelId,RowOperate rowOperate, Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		mrowDao.updateRowHidden(excelId, sheetId, alias, false);

		MRow mrow = mrowDao.getMRow(excelId, sheetId, alias);
		mrowColDao.updateRowColLength(excelId, sheetId, "rList", alias,
				mrow.getHeight());

		// mcellDao.updateHidden("rowId", alias, false, excelId);
		msheetDao.updateStep(excelId, sheetId, step);

	}

	@Override
	public void updateRowHeight( String excelId,RowHeight rowHeight,
			Integer step) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		String alias = sortRList.get(rowHeight.getRow()).getAlias();
		Integer height = rowHeight.getOffset();
		mrowDao.updateRowHeight(excelId, sheetId, alias, height);
		mrowColDao.updateRowColLength(excelId, sheetId, "rList", alias, height);
		msheetDao.updateStep(excelId, sheetId, step);

	}
	
	private void setMRowProperty(MRow mrow, Content content,
			CustomProp customProp) {
		try {
			Class<? extends Content> c = content.getClass();
			Content mc = mrow.getProps().getContent();
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
			CustomProp mp = mrow.getProps().getCustomProp();
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
	
	private void creatMCell(MRow mrow,MCol mcol,String sheetId,List<Object> tempList) {
		try {
			Content content = mcol.getProps().getContent();
			Class<? extends Content> c = content.getClass();
			Field[] cFields = c.getDeclaredFields();
			
			Content mc = mrow.getProps().getContent();
			
			for (Field f : cFields) {
				f.setAccessible(true);
				Object value = f.get(content);
				if (null != value) {
					Field mf = mc.getClass().getDeclaredField(f.getName());
					mf.setAccessible(true);
					Object mvalue = mf.get(mc);
					if(null != mvalue){
						MRowColCell mrcc = new MRowColCell();
						String id = mrow.getAlias()+"_"+mcol.getAlias();
						mrcc.setCellId(id);
						mrcc.setRow(mrow.getAlias());
						mrcc.setCol(mcol.getAlias());
						mrcc.setSheetId(sheetId);
						tempList.add(mrcc);// 关系表
						MCell mcell = new MCell(id, sheetId);
						
						// 将行、列自带的属性赋值给单元格
						OperProp rp = mrow.getProps();
						OperProp cp = mcol.getProps();
						setMCellProperty(mcell, cp.getContent(),
								cp.getBorder(), cp.getCustomProp());
						setMCellProperty(mcell, rp.getContent(),
								rp.getBorder(), rp.getCustomProp());
						break;
					}
				}
			}

			/*Class<? extends CustomProp> p = customProp.getClass();
			CustomProp mp = mrow.getProps().getCustomProp();
			Field[] pFields = p.getDeclaredFields();
			for (Field f : pFields) {
				f.setAccessible(true);
				Object value = f.get(customProp);
				if (null != value) {
					Field pf = mp.getClass().getDeclaredField(f.getName());
					pf.setAccessible(true);
					pf.set(mp, value);
				}
			}*/

		} catch (Exception e) {

			e.printStackTrace();
		}
	}
	
	private void setMCellProperty(MCell mcell, Content content, Border border,
			CustomProp customProp) {
		try {
			Class<? extends Content> c = content.getClass();
			Content mc = mcell.getContent();
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

			Class<? extends Border> b = border.getClass();
			Border mb = mcell.getBorder();
			Field[] bFields = b.getDeclaredFields();
			for (Field f : bFields) {
				f.setAccessible(true);
				Object value = f.get(border);
				if (null != value) {
					Field bf = mb.getClass().getDeclaredField(f.getName());
					bf.setAccessible(true);
					bf.set(mb, value);
				}
			}

			Class<? extends CustomProp> p = customProp.getClass();
			CustomProp mp = mcell.getCustomProp();
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

}
