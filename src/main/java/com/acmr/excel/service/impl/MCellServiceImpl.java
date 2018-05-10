package com.acmr.excel.service.impl;

import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.mongo.MExcel;
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

	
	public void saveContent(Cell cell, int step, String excelId) {
		
		int rowIndex = cell.getCoordinate().getStartRow();
		int colIndex = cell.getCoordinate().getStartCol();
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId);
		mrowColDao.getRowList(sortRcList, excelId);
		
	    MExcelCell mExcelCell = new MExcelCell();
		ExcelCell excelCell = new ExcelCell();
		mExcelCell.setExcelCell(excelCell);
	
		String content = cell.getContent();
		try {
			content = java.net.URLDecoder.decode(content, "utf-8");
			
			excelCell.setText(content);
			excelCell.setValue(content);
			CellFormateUtil.autoRecognise(content, excelCell);
			
			mExcelCell.setRowId(sortRcList.get(rowIndex).getAlias());
			mExcelCell.setColId(sortClList.get(colIndex).getAlias());
			mExcelCell.setId(mExcelCell.getRowId()+"_"+mExcelCell.getColId());
			baseDao.update(excelId, mExcelCell);
			MExcel me = new MExcel();
			me.setId(excelId);
			me.setStep(step);
			baseDao.update(excelId, me);
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		

	}

}
