package com.acmr.excel.service;

import com.acmr.excel.model.Frozen;

public interface MSheetService {

	/**
	 * 冻结
	 * @param frozen
	 * @param excelId
	 * @param step
	 */
	void frozen(Frozen frozen,String excelId,Integer step);
	/**
	 * 解除冻结
	 * @param excelId
	 * @param step
	 */
	void unFrozen(String excelId,Integer step);
	
	/**
	 * 获取步骤值
	 * @param excelId
	 * @param sheetId
	 * @return
	 */
	int getStep(String excelId,String sheetId);
}
