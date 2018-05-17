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
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.MRowService;

import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelRow;

@Service("mrowService")
public class MRowServiceImpl implements MRowService {

	@Resource
	private BaseDao baseDao;
	
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;
	
	public void insertRow(RowOperate rowOperate,String excelId){
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId);
		mrowColDao.getRowList(sortRcList, excelId);
		String alias = sortRcList.size()+1+"";
		
		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		rowCol.setLength(19);
		rowCol.setPreAlias(sortRcList.get(rowOperate.getRow()-1).getAlias());
		mrowColDao.insertRowCol(excelId, rowCol, "rList");
		//修改这行后面的前行别名
		String nextAlias = sortRcList.get(rowOperate.getRow()+1).getAlias();
		mrowColDao.updateRowCol(excelId, "rList", nextAlias, alias);
		//存一个行样式
		MExcelRow mer = new MExcelRow();
		ExcelRow er = new ExcelRow(alias);
		mer.setExcelRow(er);
		baseDao.update(excelId, mer);
		
		Integer row  = Integer.parseInt(sortRcList.get(rowOperate.getRow()).getAlias());
		List<RowColCell> relationList = mcellDao.getRowColCellList(excelId, row);
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
		String rowId = sortRcList.get(rowOperate.getRow()).getAlias();//行别名
		List<MExcelCell> sortMCellList = new ArrayList<MExcelCell>();
		
		for(RowCol rc: sortClList){
			String cellId = relationMap.get(rowId+"_"+rc.getAlias());
			sortMCellList.add(cellMap.get(cellId));
		}
		
		List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		for(int i=0;i<sortClList.size();i++){
			MExcelCell mec = sortMCellList.get(i);
			RowCol rc = sortClList.get(i);
			if(mec.getRowspan()==1){
				//new新增一行的MExcelCell对象
				MExcelCell me = new MExcelCell();
				me.setId(alias+"_"+rc.getAlias());
				me.setRowId(alias);
				me.setColId(rc.getAlias());
				me.setRowspan(1);
				me.setColspan(1);
				ExcelCell ec = new ExcelCell();
				me.setExcelCell(ec);
				tempList.add(me);
				RowColCell rcc = new RowColCell();
				rcc.setCellId(alias+"_"+rc.getAlias());
				rcc.setCol(Integer.parseInt(rc.getAlias()));
				rcc.setRow(Integer.parseInt(alias));
				tempList.add(rcc);
			}else if(mec.getRowspan()>1){
				if(mec.getRowId().equals(rowId) ){
					//new新增一行的MExcelCell对象,是以选定行作为开始行的合并单元格
					MExcelCell me = new MExcelCell();
					me.setId(alias+"_"+rc.getAlias());
					me.setRowId(alias);
					me.setColId(rc.getAlias());
					me.setRowspan(1);
					me.setColspan(1);
					ExcelCell ec = new ExcelCell();
					me.setExcelCell(ec);
					tempList.add(me);
					//关系对象
					RowColCell rcc = new RowColCell();
					rcc.setCellId(alias+"_"+rc.getAlias());
					rcc.setCol(Integer.parseInt(rc.getAlias()));
					rcc.setRow(Integer.parseInt(alias));
					tempList.add(rcc);
				}else{
					if(mec.getColId().equals(rc.getAlias())){
						mec.setRowspan(mec.getRowspan()+1);
						tempList.add(mec);
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(rc.getAlias()));
						rcc.setRow(Integer.parseInt(alias));
						tempList.add(rcc);
					}else{
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(rc.getAlias()));
						rcc.setRow(Integer.parseInt(alias));
						tempList.add(rcc);
					}
				}
			}
		}
		baseDao.update(excelId, tempList);
		
	}

	@Override
	public void delRow(RowOperate rowOperate, String excelId) {
		
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId);
		mrowColDao.getColList(sortCList, excelId);
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		//删除行
		mrowColDao.delRowCol(excelId, alias, "rList");
		//修改这行后面一行的前一行别名
		
		
	}

}
