package com.acmr.excel.service;

import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OuterPaste;
import com.acmr.excel.model.SheetParam;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.mongo.MSheet;

public interface MSheetService {

	/**
	 * 冻结
	 * 
	 * @param frozen
	 * @param excelId
	 * @param step
	 */
	void frozen(Frozen frozen, String excelId, Integer step);

	/**
	 * 解除冻结
	 * 
	 * @param excelId
	 * @param step
	 */
	void unFrozen(String excelId, Integer step);

	/**
	 * 获取步骤值
	 * 
	 * @param excelId
	 * @param sheetId
	 * @return
	 */
	int getStep(String excelId, String sheetId);

	/**
	 * 外部粘贴
	 * 
	 * @param paste
	 * @param excelId
	 * @param step
	 */
	void paste(OuterPaste paste, String excelId, Integer step);

	/**
	 * 判断是否可以进行外部粘贴粘贴
	 * 
	 * @param paste
	 * @param excelId
	 * @return
	 */
	boolean isAblePaste(OuterPaste paste, String excelId);

	/**
	 * 判断是否可以进行内部剪切
	 * 
	 * @param copy
	 * @param excelId
	 * @return
	 */
	boolean isCut(Copy copy, String excelId);
	
	/**
	 * 判断是否可以进行内部复制
	 * 
	 * @param copy
	 * @param excelId
	 * @return
	 */
	boolean isCopy(Copy copy, String excelId);

	/**
	 * 剪切复制
	 * 
	 * @param copy
	 * @param excelId
	 * @param step
	 */
	void cut(Copy copy, String excelId, Integer step);

	/**
	 * 复制
	 * 
	 * @param copy
	 * @param excelId
	 * @param step
	 */
	void copy(Copy copy, String excelId, Integer step);
	
	/**
	 * 前进
	 * @param excelId
	 */
	void redo(String excelId);
	
	/**
	 * 后退
	 * @param excelId
	 */
	void undo(String excelId);
	
	/**
	 * 更新表格保护状态
	 * @param sheetParam
	 * @param excelId
	 */
	void updateProtect(String excelId,MSheet msheet);
}
