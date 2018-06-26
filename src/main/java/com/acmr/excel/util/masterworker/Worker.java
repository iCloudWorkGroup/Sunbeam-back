package com.acmr.excel.util.masterworker;

import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.ConcurrentLinkedQueue;

import org.springframework.data.mongodb.core.MongoTemplate;

public class Worker implements Runnable {

	private ConcurrentLinkedQueue<List<Object>> workQueue;
	private ConcurrentHashMap<String, Object> resultMap;
	private MongoTemplate mongoTemplate;
	private String excelId;
	
	public void setWorkQueue(ConcurrentLinkedQueue<List<Object>> workQueue) {
		this.workQueue = workQueue;
	}

	public void setResultMap(ConcurrentHashMap<String, Object> resultMap) {
		this.resultMap = resultMap;
	}
	

	public void setMongoTemplate(MongoTemplate mongoTemplate) {
		this.mongoTemplate = mongoTemplate;
	}
	
	public void setExcelId(String excelId) {
		this.excelId = excelId;
	}

	@Override
	public void run() {
		while(true){
			List<Object> input = this.workQueue.poll();
			if(input == null) break;
			Object output = handle(input);
			//this.resultMap.put(Integer.toString(input.getId()), output);
		}
	}

	private Object handle(List<Object> tempList) {
		Object output = null;
		mongoTemplate.insert(tempList, excelId);
		return output;
	}



}
