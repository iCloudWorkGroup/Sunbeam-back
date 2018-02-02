package com.acmr.excel.dao;



import com.acmr.excel.model.OnlineExcel;
import java.sql.SQLException;
import java.util.List;

public interface ExcelDao {
	/**
	 * 保存excel
	 * 
	 * @param excel
	 *            OnlineExcel对象
	 * @throws Exception
	 */
	public void saveExcel(OnlineExcel excel) throws Exception;

	/**
	 * 获得所有的OnlineExcel对象
	 * 
	 * @return OnlineExcel对象集合
	 */
	public List<OnlineExcel> getAllExcel();

	/**
	 * 通过id获得OnlineExcel对象
	 * 
	 * @param id
	 * @return OnlineExcel
	 */
	public OnlineExcel getOnlineExcel(String id);

	/**
	 * 通过id获得JsonObject
	 * 
	 * @param id
	 * @return JsonObject
	 */
	public OnlineExcel getJsonObjectByExcelId(String excelId) throws SQLException;
	public int getByExcelId(String excelId) throws SQLException;

	/**
	 * 更新OnlineExcel
	 * 
	 * @param OnlineExcel
	 *            OnlineExcel对象
	 * @throws Exception
	 */

	public void updateExcel(OnlineExcel excel) throws Exception;
}
