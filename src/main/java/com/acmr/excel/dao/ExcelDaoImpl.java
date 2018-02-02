package com.acmr.excel.dao;




import java.sql.SQLException;
import java.util.List;

import org.springframework.jdbc.core.support.JdbcDaoSupport;

import com.acmr.excel.model.OnlineExcel;

public class ExcelDaoImpl extends JdbcDaoSupport  implements ExcelDao {
	/**
	 * 保存excel
	 * 
	 * @param excel
	 *            OnlineExcel对象
	 * @throws Exception
	 */
	@Override
	public void saveExcel(OnlineExcel excel) throws Exception {
		String sql = "insert into excel(excelId,currentTime,name) values(?,NOW(),?)";
//		LobHandler lobHandler = new DefaultLobHandler();  // reusable object
//		   this.getJdbcTemplate().execute(sql, new AbstractLobCreatingPreparedStatementCallback(lobHandler){
//			@Override
//			protected void setValues(PreparedStatement ps, LobCreator lobCreator) throws SQLException, DataAccessException {
//				ps.setString(1, excel.getExcelId());
//				lobCreator.setBlobAsBinaryStream(ps, 2, 
//						new ByteArrayInputStream(excel.getExcelObject().getBytes()), excel.getExcelObject().length());
//				ps.setString(3, excel.getName());
//			}
//			   
//		   });
		this.getJdbcTemplate().update(sql, new Object[]{excel.getExcelId(),excel.getName()});
	}

	/**
	 * 获得所有的OnlineExcel对象
	 * 
	 * @return OnlineExcel对象集合
	 */

	@Override
	public List<OnlineExcel> getAllExcel() {
		String sql = "select id,excelId,currentTime from excel order by currenttime desc";
		return this.getJdbcTemplate().query(sql, new OnlineExcel());
	}


	/**
	 * 通过id获得OnlineExcel对象
	 * 
	 * @param id
	 * @return OnlineExcel
	 */

	@Override
	public OnlineExcel getOnlineExcel(String excelId) {
		String sql = "select * from excel where excelId = ?";
		return this.getJdbcTemplate().queryForObject(sql, new Object[]{},OnlineExcel.class);
	}

	/**
	 * 通过id获得JsonObject
	 * 
	 * @param id
	 * @return JsonObject
	 */

	@Override
	public OnlineExcel getJsonObjectByExcelId(String excelId) throws SQLException {
//		String sql = "select excelObject from excel where excelId = '" + excelId +"'";
//		final LobHandler lobHandler = new DefaultLobHandler();
//		final OnlineExcel oe = new OnlineExcel();
//		 this.getJdbcTemplate().query(sql, new AbstractLobStreamingResultSetExtractor(){
//			@Override
//			protected void streamData(ResultSet rs) throws SQLException, IOException, DataAccessException {
//				byte[] b = lobHandler.getBlobAsBytes(rs, "excelObject");
//				oe.setExcelObject(new String(b,"UTF8"));
//			}
//			
//		});
//		 return oe;
		return new OnlineExcel();
	}

	/**
	 * 更新OnlineExcel
	 * 
	 * @param OnlineExcel
	 *            OnlineExcel对象
	 * @throws Exception
	 */

	@Override
	public void updateExcel(OnlineExcel excel) throws Exception {
		String sql = "update excel t set t.jsonObject = ? , t.currentTime = NOW() where t.excelId = ?";
		this.getJdbcTemplate().update(sql,
				new Object[]{excel.getExcelObject().getBytes("utf-8"),excel.getExcelId()},
				new int[]{java.sql.Types.BLOB,java.sql.Types.VARCHAR}); 

	}

	@Override
	public int getByExcelId(String excelId) throws SQLException {
		String sql = "select count(id) from excel where excelId = ?";
		return this.getJdbcTemplate().queryForObject(sql, new Object[]{excelId},Integer.class);
	}
}
