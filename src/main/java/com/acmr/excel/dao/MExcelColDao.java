package com.acmr.excel.dao;

import com.acmr.excel.model.mongo.MExcelColumn;

public interface MExcelColDao {

/**
 * 根据列样式中的code值，删除列样式	
 * @param excelId
 * @param alias
 */
 void delExcelCol(String excelId,String alias);
 /**
  * 根据列样式中的code值，修改列样式的隐藏属性
  * @param excelId
  * @param alias
  * @param status
  */
 void updateColHiddenStatus(String excelId,String alias,boolean status);
 
 /**
  * 根据列样式中的code，更新列宽度
  * @param excelId
  * @param alias
  * @param width
  */
 void updateColWidth(String excelId,String alias,Integer width);
 
 MExcelColumn getMExcelCol(String excelId,String alias);
}
