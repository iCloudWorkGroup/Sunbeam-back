package com.acmr.excel.dao.base;

import java.util.Collection;

import org.springframework.data.mongodb.core.MongoTemplate;

public class ThreadDao implements Runnable{
	
	private MongoTemplate mongoTemplate;
	private String excelId;
	private Collection<? extends Object> object;
	

	@Override
	public void run() {
		
		mongoTemplate.insert(object, excelId);
	}

	
	
	public ThreadDao(){}

	public ThreadDao(MongoTemplate mongoTemplate, String excelId,
			Collection<? extends Object> object) {
		
		this.mongoTemplate = mongoTemplate;
		this.excelId = excelId;
		this.object = object;
	}
	

}
