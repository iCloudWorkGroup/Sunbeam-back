package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.MColService;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;

@Service("mcolService")
public class MColServiceImpl implements MColService {
	
	@Resource
	private BaseDao baseDao;
	
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;

	@Override
	public void insertCol(ColOperate colOperate, String excelId) {
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId);
		mrowColDao.getRowList(sortRcList, excelId);
		String alias = sortClList.size()+1+"";
		
		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		rowCol.setLength(19);
		rowCol.setPreAlias(sortRcList.get(colOperate.getCol()-1).getAlias());
		mrowColDao.insertRowCol(excelId, rowCol, "cList");
		//修改这列后面的前列别名
		String nextAlias = sortRcList.get(colOperate.getCol()+1).getAlias();
		mrowColDao.updateRowCol(excelId, "cList", nextAlias, alias);
		//存一个列样式
		MExcelColumn meco = new MExcelColumn();
		ExcelColumn eco = new ExcelColumn(alias);
		meco.setExcelColumn(eco);
		baseDao.update(excelId, meco);
		
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
			if(mec.getColspan()==1){
				//new新增一行的MExcelCell对象
				MExcelCell me = new MExcelCell();
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
				tempList.add(rcc);
			}else if(mec.getColspan()>1){
				if(mec.getColId().equals(colId) ){
					//new新增一行的MExcelCell对象,是以选定列作为开始列的合并单元格
					MExcelCell me = new MExcelCell();
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
					tempList.add(rcc);
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
		baseDao.update(excelId, tempList);
		
	}

	

}
