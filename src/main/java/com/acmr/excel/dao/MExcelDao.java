package com.acmr.excel.dao;

import com.acmr.excel.model.mongo.MExcel;

public interface MExcelDao {
 /**
  * 更新步骤
  * @param excelId
  * @param step
  */
  void 	updateStep(String excelId,Integer step);
  /**
   * 更新冻结属性
   * @param mexcel
   * @param excelId
   */
  void updateFrozen(MExcel mexcel,String excelId);
  
  /**
   * 解除冻结
   * @param mexcel
   * @param excelId
   */
  void updateUnFrozen(MExcel mexcel,String excelId);
  
  void updateMaxRow(Integer alias,String excelId);
  
  void updateMaxCol(Integer alias,String excelId);
  MExcel getMExcel(String excelId);
}
