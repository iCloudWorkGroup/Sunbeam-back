package com.acmr.excel.service;

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

	CompleteExcel reload(String excelId, String sheetId, Integer rowBegin,
			Integer rowEnd, Integer colBegin, Integer colEnd);

}
