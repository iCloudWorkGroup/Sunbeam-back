package com.acmr.excel.service;

import com.acmr.excel.model.complete.rows.RowOperate;

public interface MRowService {
	
  /**
   * 插入一行	
   * @param rowOperate
   */
  void insertRow(RowOperate rowOperate,String excelId);
  
  void delRow(RowOperate rowOperate,String excelId);
}
