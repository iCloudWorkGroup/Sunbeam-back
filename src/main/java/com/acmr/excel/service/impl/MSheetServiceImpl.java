package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MExcelDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.mongo.MExcel;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.MSheetService;
@Service("msheetService")
public class MSheetServiceImpl implements MSheetService {
	@Resource
	private BaseDao baseDao;
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MExcelDao mexcelDao;
	
	@Override
	public void frozen(Frozen frozen, String excelId, Integer step) {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId);
		mrowColDao.getColList(sortCList, excelId);
		String oprRow = sortRList.get(frozen.getOprRow()).getAlias();
		String oprCol = sortCList.get(frozen.getOprCol()).getAlias();
		String viewRow = sortRList.get(frozen.getViewRow()).getAlias();
		String viewCol = sortCList.get(frozen.getViewCol()).getAlias();
		MExcel mexcel = new MExcel();
		mexcel.setFreeze(true);
		mexcel.setRowAlias(oprRow);
		mexcel.setColAlias(oprCol);
		mexcel.setViewRowAlias(viewRow);
		mexcel.setViewColAlias(viewCol);
		mexcel.setStep(step);
		
		mexcelDao.updateFrozen(mexcel, excelId);

	}

	@Override
	public void unFrozen(String excelId, Integer step) {
		MExcel mexcel = new MExcel();
		mexcel.setFreeze(false);
		mexcel.setStep(step);
		mexcel.setId(excelId);
		mexcelDao.updateUnFrozen(mexcel, excelId);
		
	}

}
