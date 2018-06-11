package com.acmr.excel.dao.impl;

import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.mongo.MRow;

@Repository("mrowDao")
public class MRowDaoImpl implements MRowDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public List<MRow> getMRowList(String excelId, String sheetId, List<String> rowList) {
		List<MRow> mRowList = mongoTemplate.find(new Query(Criteria.where("alias").in(rowList).and("sheetId").is(sheetId).and("_class").is(MRow.class.getName())), MRow.class, excelId);
		return mRowList;
	}

	@Override
	public void delMRow(String excelId, String sheetId, String alias) {
		
		mongoTemplate.remove(new Query(Criteria.where("alias").is(alias).and("_class").is(MRow.class.getName()).and("sheetId").is(sheetId)), 
				MRow.class, excelId);
		
	}

}
