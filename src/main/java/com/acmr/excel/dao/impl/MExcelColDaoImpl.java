package com.acmr.excel.dao.impl;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MExcelColDao;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;

@Repository("mexcelColDao")
public class MExcelColDaoImpl implements MExcelColDao {
	
	@Resource
	private MongoTemplate mongoTemplate;

	@Override
	public void delExcelCol(String excelId, String alias) {
		
	mongoTemplate.remove(new Query(Criteria.where("excelColumn.code").is(alias)), MExcelColumn.class, excelId);
   }

	@Override
	public void updateColHiddenStatus(String excelId, String alias, boolean status) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("excelColumn.code").is(alias));
		Update update = new Update();
		update.set("excelColumn.columnhidden", status);
		mongoTemplate.updateFirst(query, update,MExcelColumn.class, excelId);
		
	}

	@Override
	public void updateColWidth(String excelId, String alias, Integer width) {
		Query query = new Query();
		query.addCriteria(Criteria.where("excelColumn.code").is(alias));
		Update update = new Update();
		update.set("excelColumn.width", width);
		mongoTemplate.updateFirst(query, update,MExcelColumn.class, excelId);
	}

	@Override
	public MExcelColumn getMExcelCol(String excelId, String alias) {
		MExcelColumn mexcelCol =  mongoTemplate.findOne(new Query(Criteria.where("excelColumn.code").is(alias)), MExcelColumn.class,excelId);
		return mexcelCol;
	}

}
