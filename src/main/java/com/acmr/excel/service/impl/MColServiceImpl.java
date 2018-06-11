package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MExcelColDao;
import com.acmr.excel.dao.MExcelDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.mongo.MExcel;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.MColService;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;

@Service("mcolService")
public class MColServiceImpl implements MColService {
	
	@Resource
	private BaseDao baseDao;
	
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MExcelColDao mexcelColDao;
	@Resource
	private MExcelDao mexcelDao;

	@Override
	public void insertCol(ColOperate colOperate, String excelId,Integer step) {
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId);
		mrowColDao.getColList(sortRcList, excelId);
		if(colOperate.getCol()>sortClList.size()-1){
			return;
		}
		
		MExcel mexcel = mexcelDao.getMExcel(excelId);
		String alias = mexcel.getMaxcol()+1+"";
		
		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		rowCol.setLength(19);
		if(colOperate.getCol() == 0){
			rowCol.setPreAlias(null);
		}else{
			rowCol.setPreAlias(sortRcList.get(colOperate.getCol()-1).getAlias());
		}
		
		mrowColDao.insertRowCol(excelId, rowCol, "cList");
		//修改这列前列别名
		String nextAlias = sortRcList.get(colOperate.getCol()).getAlias();
		mrowColDao.updateRowCol(excelId, "cList", nextAlias, alias);
		//存一个列样式
		MExcelColumn meco = new MExcelColumn();
		ExcelColumn eco = new ExcelColumn(alias);
		meco.setExcelColumn(eco);
		baseDao.update(excelId, meco);
		
		mexcelDao.updateMaxCol(mexcel.getMaxcol()+1, excelId);
		
		Integer col  = Integer.parseInt(sortClList.get(colOperate.getCol()).getAlias());
		List<RowColCell> relationList = mcellDao.getRowColCellList(col, excelId);
		List<String> cellIdList = new ArrayList<>();
		Map<String,String> relationMap = new HashMap<String,String>();
		for(RowColCell rcc:relationList){
			cellIdList.add(rcc.getCellId());
			relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
		}
		List<MExcelCell> mcellList = mcellDao.getMCellList(excelId, cellIdList);
		Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
		for(MExcelCell mec:mcellList){
			cellMap.put(mec.getId(), mec);
		}
		String colId = sortClList.get(colOperate.getCol()).getAlias();//列别名
		List<MExcelCell> sortMCellList = new ArrayList<MExcelCell>();
		
		for(RowCol rc: sortRcList){
			String cellId = relationMap.get(rc.getAlias()+"_"+colId);
			sortMCellList.add(cellMap.get(cellId));
		}
		
		List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		for(int i=0;i<sortRcList.size();i++){
			MExcelCell mec = sortMCellList.get(i);
			RowCol rc = sortRcList.get(i);
			if(null == mec){
				continue;
			}
			if(null!=mec&&mec.getColspan()==1){
				//new新增一行的MExcelCell对象
			/*	MExcelCell me = new MExcelCell();
				me.setId(rc.getAlias()+"_"+alias);
				me.setRowId(rc.getAlias());
				me.setColId(alias);
				me.setRowspan(1);
				me.setColspan(1);
				ExcelCell ec = new ExcelCell();
				me.setExcelCell(ec);
				tempList.add(me);
				RowColCell rcc = new RowColCell();
				rcc.setCellId(rc.getAlias()+"_"+alias);
				rcc.setCol(Integer.parseInt(alias));
				rcc.setRow(Integer.parseInt(rc.getAlias()));
				tempList.add(rcc);*/
			}else if(mec.getColspan()>1){
				if(mec.getColId().equals(colId) ){
					//new新增一行的MExcelCell对象,是以选定列作为开始列的合并单元格
				/*	MExcelCell me = new MExcelCell();
					me.setId(rc.getAlias()+"_"+alias);
					me.setRowId(rc.getAlias());
					me.setColId(alias);
					me.setRowspan(1);
					me.setColspan(1);
					ExcelCell ec = new ExcelCell();
					me.setExcelCell(ec);
					tempList.add(me);
					//关系对象
					RowColCell rcc = new RowColCell();
					rcc.setCellId(rc.getAlias()+"_"+alias);
					rcc.setCol(Integer.parseInt(alias));
					rcc.setRow(Integer.parseInt(rc.getAlias()));
					tempList.add(rcc);*/
				}else{
					if(mec.getRowId().equals(rc.getAlias())){
						mec.setColspan(mec.getColspan()+1);
						tempList.add(mec);
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(alias));
						rcc.setRow(Integer.parseInt(rc.getAlias()));
						tempList.add(rcc);
					}else{
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(alias));
						rcc.setRow(Integer.parseInt(rc.getAlias()));
						tempList.add(rcc);
					}
				}
			}
		}
		if(tempList.size()>0){
			baseDao.update(excelId, tempList);
		}
		mexcelDao.updateStep(excelId, step);
	}

	@Override
	public void delCol(ColOperate colOperate, String excelId,Integer step) {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortRList, excelId);
		mrowColDao.getColList(sortCList, excelId);
		if(colOperate.getCol()>sortCList.size()-1){
			return;
		}
		//列别名
		String alias = sortCList.get(colOperate.getCol()).getAlias();
		//删除列
		mrowColDao.delRowCol(excelId, alias, "cList");
		//修改这列后面一列的前列别名
		int index = colOperate.getCol();
		String frontAlias;
		if(colOperate.getCol()==0){
			frontAlias = null;
		}else{
			frontAlias = sortCList.get(index-1).getAlias();//前一列别名
		}
		 
		String backAlias = sortCList.get(index+1).getAlias();//后一列别名
		mrowColDao.updateRowCol(excelId, "cList", backAlias, frontAlias);
		//删除列样式记录
		mexcelColDao.delExcelCol(excelId, alias);
	    List<RowColCell> relationList = mcellDao.getRowColCellList(Integer.parseInt(alias),excelId);
	    List<String> cellIdList = new ArrayList<String>();
	    Map<String,String> relationMap = new HashMap<String,String>();
	    for(RowColCell rcc:relationList){
	    	
	    	relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
	    	cellIdList.add(rcc.getCellId());
	    }
        List<MExcelCell> cellList = mcellDao.getMCellList(excelId, cellIdList);	
        List<MExcelCell> sortMcellList = new ArrayList<MExcelCell>();
        Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
        for(MExcelCell ec:cellList){
        	cellMap.put(ec.getId(), ec);
        }
		for(RowCol rc:sortRList){
			String cellId = relationMap.get(rc.getAlias()+"_"+alias);
			sortMcellList.add(cellMap.get(cellId));
		}
		//删除关系表
		mcellDao.delRowColCell(excelId,"col", Integer.parseInt(alias));
		cellIdList.clear();//存需要删除的MExcelCell的Id
		cellList.clear();//存需要修改或增加的MExcelCell对象
		for(int i=0;i<sortRList.size();i++){
			String row = sortRList.get(i).getAlias();
			MExcelCell ec = sortMcellList.get(i);
			if(null == ec){
				continue;
			}
			if(ec.getColspan() == 1){
				cellIdList.add(ec.getId());
			}else{
				if(ec.getColId().equals(alias)){
					if(ec.getRowId().equals(row)){
						//删除老的MExcelCell
						cellIdList.add(ec.getId());
						MExcelCell mec = ec;
						mec.setColId(backAlias);
						String id = row+"_"+backAlias;
						//修改合并单元格其他关系表的cellId
						mcellDao.updateRowColCell(excelId, mec.getId(), id);
						mec.setId(id);
						mec.setColspan(ec.getColspan()-1);
						
						ExcelCell cell = ec.getExcelCell();
						cell.setColspan(cell.getColspan()-1);
						mec.setExcelCell(cell);
						cellList.add(mec);//插入新MExcelCell
								
					}
				}else{
					if(ec.getRowId().equals(row)){
						MExcelCell mec = ec;
						mec.setColspan(mec.getColspan()-1);
						cellList.add(mec);
					}
				} 
			}
			
		}
		
		mcellDao.delCell(excelId, cellIdList);
		if(cellList.size()>0){
		 baseDao.update(excelId, cellList);
		}
		mexcelDao.updateStep(excelId, step);
		
	}

	@Override
	public void hideCol(ColOperate colOperate, String excelId,Integer step) {
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId);
		String alias = sortCList.get(colOperate.getCol()).getAlias();
		
		mexcelColDao.updateColHiddenStatus(excelId, alias, true);
		
		mrowColDao.updateRowColLength(excelId, "cList", alias, 0);
		//mcellDao.updateHidden("colId", alias, true, excelId);
		mexcelDao.updateStep(excelId, step);
	}

	@Override
	public void showCol(ColOperate colOperate, String excelId,Integer step) {
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId);
		String alias = sortCList.get(colOperate.getCol()).getAlias();
		mexcelColDao.updateColHiddenStatus(excelId, alias, false);
		
		MExcelColumn mcol = mexcelColDao.getMExcelCol(excelId, alias);
		mrowColDao.updateRowColLength(excelId, "cList", alias, mcol.getExcelColumn().getWidth());
		//mcellDao.updateHidden("colId", alias, false, excelId);
		mexcelDao.updateStep(excelId, step);
	}

	@Override
	public void updateColWidth(ColWidth colWidth, String excelId, Integer step) {
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortCList, excelId);
		String alias = sortCList.get(colWidth.getCol()).getAlias();
		Integer width = colWidth.getOffset();
		
		mexcelColDao.updateColWidth(excelId, alias, width);
		mrowColDao.updateRowColLength(excelId, "cList", alias, width);
		
		mexcelDao.updateStep(excelId, step);
		
	}

	

}
