package com.acmr.excel.service;

import com.acmr.excel.model.complete.rows.ColOperate;

public interface MColService {
	
 /**
  * 插入一列	
  * @param colOperate
  * @param excelId 数据集
  */ 
 void insertCol(ColOperate colOperate,String excelId);

}
