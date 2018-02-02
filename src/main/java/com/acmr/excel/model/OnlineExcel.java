package com.acmr.excel.model;

import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.springframework.jdbc.core.RowMapper;

public class OnlineExcel implements RowMapper<OnlineExcel> {
	private long id;
	private String excelId;
	private String excelObject;
	private Date currentTime;
	private String name;

	public String getName() {
		return name;
	}

	public void setName(String name) {
		this.name = name;
	}

	public long getId() {
		return id;
	}

	public void setId(long id) {
		this.id = id;
	}

	public String getExcelId() {
		return excelId;
	}

	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	public String getExcelObject() {
		return excelObject;
	}

	public void setExcelObject(String excelObject) {
		this.excelObject = excelObject;
	}

	public Date getCurrentTime() {
		return currentTime;
	}

	public void setCurrentTime(Date currentTime) {
		this.currentTime = currentTime;
	}

	@Override
	public OnlineExcel mapRow(ResultSet rs, int rowNum) throws SQLException {
		OnlineExcel oe = new OnlineExcel();
		oe.setId(rs.getLong("id"));
		oe.setExcelId(rs.getString("excelId"));
		//oe.setExcelObject(rs.getNString("excelObject"));
		//oe.setName(rs.getString("name"));
		String currentTime = rs.getString("currentTime");
		SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
		try {
			oe.setCurrentTime(sdf.parse(currentTime));
		} catch (ParseException e) {
			e.printStackTrace();
		}
		return oe;
	}
}
