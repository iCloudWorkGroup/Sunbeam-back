package com.acmr.excel.service;

import java.util.List;

import com.acmr.excel.model.complete.CompleteExcel;

import acmr.excel.pojo.ExcelBook;

public interface MBookService {
	/**
	 * 保存上传过来的Excel数据
	 * 
	 * @param book
	 * @param excelId
	 * @return
	 */
	boolean saveExcelBook(ExcelBook book, String excelId);

	/**
	 * 用于重新加载及滚动
	 * @param excelId
	 * @param sheetId
	 * @param rowBegin
	 * @param rowEnd
	 * @param colBegin
	 * @param colEnd
	 * @param type  0 重新加载 1滚动
	 * @return
	 */
	CompleteExcel reload(String excelId, String sheetId, Integer rowBegin,
			Integer rowEnd, Integer colBegin, Integer colEnd,int type);
	
	/**
	 * 根据excelId,返回一个ExcelBook对象的表格
	 * @param excelId
	 * @param step
	 * @return
	 */
	ExcelBook reloadExcelBook(String excelId,Integer step);
	
	/**
	 * 获取所有的表明
	 * @return
	 */
	List<String> getExcels();

}
