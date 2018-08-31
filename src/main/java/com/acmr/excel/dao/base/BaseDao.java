package com.acmr.excel.dao.base;

import java.util.Collection;
import java.util.Collections;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import javax.annotation.Resource;

import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.BasicQuery;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.stereotype.Repository;

import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.mongodb.BasicDBObject;
import com.mongodb.DBObject;

@Repository("baseDao")
public class BaseDao {

	@Resource
	private MongoTemplate mongoTemplate;

	public boolean update(String excelId, Object object) {
		if (!object.equals(null)) {
			mongoTemplate.save(object, excelId);
		}
		return true;
	}

	public boolean insert(String excelId, Object object) {

		mongoTemplate.insert(object, excelId);
		return true;
	}

	public boolean insertList(String excelId, Collection<? extends Object> object) {
		object.removeAll(Collections.singleton(null));
		mongoTemplate.insert(object, excelId);
		return true;
	}

	public boolean del(String excelId, Object object) {
		if (object.getClass() == MRow.class) {
			MRow mrow = (MRow) object;
			Query query = new Query();
			query.addCriteria(Criteria.where("alias").is(mrow.getAlias())
					.and("sheetId").is(mrow.getSheetId()).and("_class")
					.is(MRow.class.getName()));
			mongoTemplate.remove(query, excelId);
		} else if (object.getClass() == MRowColCell.class) {
			MRowColCell mrcc = (MRowColCell) object;
			Query query = new Query();
			query.addCriteria(Criteria.where("row").is(mrcc.getRow()).and("col")
					.is(mrcc.getCol()).and("sheetId").is(mrcc.getSheetId())
					.and("_class").is(MRowColCell.class.getName()));
			mongoTemplate.remove(query, excelId);
		} else if (object.getClass() == MCol.class) {
			MCol mcol = (MCol) object;
			Query query = new Query();
			query.addCriteria(Criteria.where("alias").is(mcol.getAlias())
					.and("sheetId").is(mcol.getSheetId()).and("_class")
					.is(MCol.class.getName()));
			mongoTemplate.remove(query, excelId);
		} else {
			mongoTemplate.remove(object, excelId);
		}
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
	
	@SuppressWarnings({ "unchecked", "rawtypes" })
	public void threadInsert(String excelId,Collection<? extends Object>  list){
		long start = System.currentTimeMillis();
		int num;
		if(list.size()/10000>0){
			num = list.size()/10000+1;
		}else{
			num = list.size()/10000;
			if(num ==0){
				num=1;
			}
		}
		
		ThreadPoolExecutor threadPool = (ThreadPoolExecutor) Executors
				.newFixedThreadPool(num);
		for(int i=0;i<num;i++){
			List<Object> li=null;
			if(i==num-1){
			   li =((List) list).subList(i*10000, list.size());
			}else{
			   li = ((List) list).subList(10000*i, 10000*(i+1));
			}
			
			threadPool.execute(new ThreadDao(mongoTemplate,excelId,li));
					
		}
		threadPool.shutdown();
		while(true){
			   if(threadPool.isTerminated()){
	                //System.out.println("Finally do something ");
	                long end = System.currentTimeMillis();
	                System.out.println("保存用时: " + (end - start) + "ms");
	                break;
	            }
		}
	}

}
