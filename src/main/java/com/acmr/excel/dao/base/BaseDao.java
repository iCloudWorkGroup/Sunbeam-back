package com.acmr.excel.dao.base;

import java.util.Collection;
import java.util.List;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.BasicQuery;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.model.mongo.MSheet;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;

@Repository("baseDao")
public class BaseDao {
	
	@Resource
	private MongoTemplate mongoTemplate;
	
	public boolean update(String excelId, Object object) {
		if(!object.equals(null)){
		 mongoTemplate.save(object, excelId);
		}
		return true;
	}
	
	public boolean insert(String excelId, Object object) {
	
		mongoTemplate.insert(object, excelId);
		return true;
	}
	
	public boolean insert(String excelId, Collection<? extends Object> object) {
		
		mongoTemplate.insert(object,excelId);;
		return true;
	}
	
	public int getStep(String id) {
		DBObject dbObject = new BasicDBObject();
		dbObject.put("_id", id);
		BasicDBObject fieldsObject = new BasicDBObject();
		fieldsObject.put("step", true);
		Query query = new BasicQuery(dbObject, fieldsObject);
		List<MSheet> result = mongoTemplate.find(query, MSheet.class, id);
		return result.get(0).getStep();
	}

}
