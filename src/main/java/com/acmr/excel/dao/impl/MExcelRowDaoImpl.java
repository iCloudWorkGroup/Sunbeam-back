package com.acmr.excel.dao.impl;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.dao.MExcelRowDao;
import com.acmr.excel.model.mongo.MExcelRow;
import com.mongodb.WriteResult;

@Repository("mexcelRowDao")
public class MExcelRowDaoImpl implements MExcelRowDao {
	
    @Resource
	private MongoTemplate mongoTemplate;
	public void delExcelRow(String excelId, String alais) {
		
	 mongoTemplate.remove(new Query(Criteria.where("excelRow.code").is(alais)), MExcelRow.class, excelId);
    
	}

}
