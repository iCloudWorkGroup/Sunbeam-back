package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MExcelCell;

@Repository("mcellDao")
public class MCellDaoImpl implements MCellDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<RowColCell> getRowColCellList(String excelId,Integer row) {
		
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("row").is(row));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		return list;
	}

	@Override
	public List<MExcelCell> getMCellList(String excelId, List<String> cellIdList) {
		
		List<MExcelCell> list = mongoTemplate.find(new Query(Criteria.where("_id").in(cellIdList)), MExcelCell.class,excelId);
		return list;
	}

	@Override
	public List<RowColCell> getRowColCellList(Integer col, String excelId) {
		
		Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("col").is(col));
		List<RowColCell> list = mongoTemplate.find(new Query(criatira), RowColCell.class,excelId);
		return list;
	}

}
