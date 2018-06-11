package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MColDao;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MExcelColumn;

@Repository("mcolDao")
public class MColDaoImpl implements MColDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MCol> getMColList(String excelId, String sheetId, List<String> colList) {
		List<MCol> mColList = mongoTemplate.find(new Query(Criteria.where("alias").in(colList).and("sheetId").is(sheetId).and("_class").is(MCol.class.getName())), MCol.class, excelId);
		return mColList;
	}

	@Override
	public void delExcelCol(String excelId, String sheetId, String alias) {
		
		mongoTemplate.remove(new Query(Criteria.where("alias").is(alias).and("_class").is(MCol.class.getName()).and("sheetId").is(sheetId)),
				MCol.class, excelId);
		
	}

}
