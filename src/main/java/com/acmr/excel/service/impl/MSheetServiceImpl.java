package com.acmr.excel.service.impl;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Frozen;
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
	
	
	@Override
	public void frozen(Frozen frozen, String excelId, Integer step) {
		String sheetId = excelId+0;
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
		
		msheetDao.updateFrozen(msheet, excelId,sheetId);

	}

	@Override
	public void unFrozen(String excelId, Integer step) {
		String sheetId = excelId+0;
		MSheet mexcel = new MSheet();
		mexcel.setFreeze(false);
		mexcel.setStep(step);
		mexcel.setId(excelId);
		msheetDao.updateUnFrozen(mexcel, excelId,sheetId);
		
	}

	@Override
	public int getStep(String excelId, String sheetId) {
		
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		
		return msheet.getStep();
	}

}
