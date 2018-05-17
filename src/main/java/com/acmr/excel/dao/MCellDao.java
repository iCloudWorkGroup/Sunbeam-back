package com.acmr.excel.dao;

import java.util.List;

import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MExcelCell;

public interface MCellDao {

 /**
  * 根据row号，查询对应的映射表
  * @param excelId  数据集
  * @param row      行号
  * @return 
  */
 List<RowColCell> getRowColCellList(String excelId,Integer row);
 
 /**
  * 根据列号，查询对应的映射表
  * @param col  列号
  * @param excelId 数据集
  * @return 
  */
 List<RowColCell> getRowColCellList(Integer col,String excelId);
 
 /**
  * 根据cellID 查询cell列表
  * @param excelId
  * @param cellIdList
  * @return
  */
 List<MExcelCell> getMCellList(String excelId,List<String> cellIdList);
}
