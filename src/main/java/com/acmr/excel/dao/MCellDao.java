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
  * 根据行别名删除关系表
  * @param excelId 数据集
  * @param field   字段名 row或col
  * @param alais   行或列的别名
  */
 void delRowColCell(String excelId,String field,Integer alias);
 
 /**
  * 根据列别名删除关系表
  * @param col     列别名
  * @param excelId 数据集
  */
 void delRowColCell(Integer col,String excelId);
 
 /**
  * 批量根据cellId修改关系表的cellId
  * @param excelId
  * @param originalId
  * @param modifiedId
  */
 void updateRowColCell(String excelId,String originalId,String modifiedId);
 
 /**
  * 根据cellID 查询cell列表
  * @param excelId
  * @param cellIdList
  * @return
  */
 List<MExcelCell> getMCellList(String excelId,List<String> cellIdList);
 
 MExcelCell getMCell(String excelId,String id);
 
 /**
  * 根据cellId删除对应MExcelCell对象
  * @param excelId    数据集
  * @param cellIdList Id集合
  */
 void delCell(String excelId,List<String> cellIdList);
 
 void updateFontname(String fontName,List<String> idList,String excelId);
 
 void updateFontSize(short size,List<String> idList,String excelId);
 
 
}
