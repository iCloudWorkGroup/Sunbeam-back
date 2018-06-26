package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.mongo.MRowColCell;

public interface MRowColCellDao {

	List<MRowColCell> getMRowColCellList(String excelId, String sheetId,
			List<String> rowList, List<String> colList);

}
