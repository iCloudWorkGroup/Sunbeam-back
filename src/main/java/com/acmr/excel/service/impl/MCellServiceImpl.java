package com.acmr.excel.service.impl;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MExcelDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.MCellService;
import com.acmr.excel.util.CellFormateUtil;

import acmr.excel.pojo.ExcelCell;

@Service("mcellService")
public class MCellServiceImpl implements MCellService {
	
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private BaseDao baseDao;
	@Resource
	private MExcelDao mexcelDao;
	@Resource
	private MCellDao mcellDao;

	
	public void saveContent(Cell cell, int step, String excelId) {
		
		int rowIndex = cell.getCoordinate().getStartRow();
		int colIndex = cell.getCoordinate().getStartCol();
		/*int rowEnd = cell.getCoordinate().getEndRow();
		int colEnd = cell.getCoordinate().getEndCol();*/
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId);
		mrowColDao.getRowList(sortRcList, excelId);
		String rowAlias = sortRcList.get(rowIndex).getAlias();
		String colAlias = sortClList.get(colIndex).getAlias();
		String id = rowAlias+"_"+colAlias;
		MExcelCell mExcelCell = mcellDao.getMCell(excelId, id);
		ExcelCell excelCell;
		if(null == mExcelCell){
			mExcelCell = new MExcelCell();
			mExcelCell.setRowspan(1);
			mExcelCell.setColspan(1);
			mExcelCell.setRowId(rowAlias);
			mExcelCell.setColId(colAlias);
			mExcelCell.setId(id);
			excelCell = new ExcelCell();
			excelCell.getCellstyle().setAlign((short)2);
			mExcelCell.setExcelCell(excelCell);
			
			RowColCell rowColCell = new RowColCell();
			rowColCell.setRow(Integer.parseInt(rowAlias));
			rowColCell.setCol(Integer.parseInt(colAlias));
			rowColCell.setCellId(id);
			baseDao.update(excelId, rowColCell);//存关系映射表
		}else{
			excelCell = mExcelCell.getExcelCell();
		}
	
		String content = cell.getContent();
		try {
			content = java.net.URLDecoder.decode(content, "utf-8");
			
			excelCell.setText(content);
			excelCell.setValue(content);
			CellFormateUtil.autoRecognise(content, excelCell);
			
			
			baseDao.update(excelId, mExcelCell);//存cell对象
			/*List<RowColCell> relationList = new ArrayList<RowColCell>();
			for(int i =rowIndex;i<=rowEnd;i++){
				for(int j=colIndex;j<=colEnd;j++){
					RowColCell rowColCell = new RowColCell();
					rowColCell.setRow(Integer.parseInt(sortRcList.get(i).getAlias()));
					rowColCell.setCol(Integer.parseInt(sortClList.get(j).getAlias()));
					rowColCell.setCellId(id);
					relationList.add(rowColCell);
				}
			}
			baseDao.update(excelId, relationList);//存关系映射表
*/			
			mexcelDao.updateStep(excelId, step);
			
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		

	}


	@Override
	public void updateFontFamily(Cell cell, int step, String excelId) {
		// TODO Auto-generated method stub
		
	}

}
