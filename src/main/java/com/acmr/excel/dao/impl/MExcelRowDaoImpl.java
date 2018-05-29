package com.acmr.excel.dao.impl;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.CriteriaDefinition;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MExcelRowDao;
import com.acmr.excel.model.mongo.MExcelRow;
import com.mongodb.WriteResult;

@Repository("mexcelRowDao")
public class MExcelRowDaoImpl implements MExcelRowDao {
	
    @Resource
	private MongoTemplate mongoTemplate;
	public void delExcelRow(String excelId, String alias) {
		
	 mongoTemplate.remove(new Query(Criteria.where("excelRow.code").is(alias)), MExcelRow.class, excelId);
    
	}
	
	public void updateRowHidden(String excelId,String alias, boolean status) {
		
		Query query = new Query();
		query.addCriteria(Criteria.where("excelRow.code").is(alias));
		Update update = new Update();
		update.set("excelRow.rowhidden", status);
		mongoTemplate.updateFirst(query, update,MExcelRow.class, excelId);
		
	}

	@Override
	public void updateRowHeight(String excelId, String alias, Integer height) {
		Query query = new Query();
		query.addCriteria(Criteria.where("excelRow.code").is(alias));
		Update update = new Update();
		update.set("height", height);
		mongoTemplate.updateFirst(query, update,MExcelRow.class, excelId);
		
	}

}
