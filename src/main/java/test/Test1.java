package test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;


public class Test1 {
	
	

	public static void main(String[] args) {
		
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		//ApplicationContext app = new ClassPathXmlApplicationContext("/config/applicationContext-core.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		//BaseDao base  =(BaseDao) applicationContext.getBean("baseDao");
		/*Criteria criatira = new Criteria();
		criatira.andOperator(Criteria.where("_id").is("cList"),
						Criteria.where("rcList.alias").is("1"));
		mongoTemplate.remove(new Query(criatira),"3bc2d656-b315-496b-aaf7-e5813c383a9c")*/;
		
		/*Query query = Query.query(Criteria.where("_id").is("rList"));
		BasicDBObject s = new BasicDBObject();
		s.put("alias","1");
		Update update = new Update();
		update.pull("rcList", s);
	    mongoTemplate.updateFirst(query, update, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");*/
		
		/*Query query = new Query();
		query.addCriteria(Criteria.where("_id").is("rList"));
		RowCol rowCol = new RowCol();
		rowCol.setAlias("1");
		rowCol.setLength(18);
		Update update = new Update();
		update.addToSet("rcList",rowCol);
		mongoTemplate.upsert(query, update, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");*/
		/*List<String> list = new ArrayList<String>();
		list.set(2, "a");*/
	/*	List<RowCol> list = new ArrayList<RowCol>();
		RowCol rc = new RowCol();
		rc.setAlias("1");
		list.add(rc);
		RowCol rc1 = new RowCol();
		rc1.setAlias("1");*/
		/*Query query = new Query();
		query.addCriteria(Criteria.where("_id").is("rList").and("rcList.alias").is("2"));
		Update update = new Update();
		update.set("rcList.$.preAlias", "1");
		mongoTemplate.updateFirst(query, update, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");*/
		
		//mongoTemplate.remove(new Query(Criteria.where("excelRow.code").is("1")), MExcelRow.class, "d02b0e50-02f0-4c7e-b2ac-486e6e487e7c");
		/*RowColCell cell = new RowColCell();
		cell.setRow(100);
		mongoTemplate.save(cell, "9dad6428-2b6d-42dd-9e9a-185fdcfc0e12");
		mongoTemplate.remove(cell, "9dad6428-2b6d-42dd-9e9a-185fdcfc0e12");
		*/
		List<Object> list = new ArrayList<Object>();
		//list.add("das");
		MRow row  = new MRow();
		MRow rwo1 = new MRow();
		row.setAlias("3");
		row.setSheetId("23");
		rwo1.setSheetId("23");
		
		list.add(row);
		list.add(rwo1);
		list.remove(row);
		
		/*Content content = row.getProps().getContent();
		
		MCol mcol = new MCol();
		mcol.setAlias("7");
		MCell mc = new MCell();
		mc.setId("1_2");
		mc.setColspan(2);
		mc.setRowspan(6);
		
		//list.add(mcol);
		//mongoTemplate.save(mc, "aa");
		//mongoTemplate.dropCollection("aa");
		Map<String,String> map = new HashMap<String,String>();
		map.put("1", "aa");
		String a = map.get("a");
		String b = map.get("1");*/
		/*Query query = new Query();
	    query.addCriteria(Criteria.where("alias").is("7"));
	    Update update = new Update();
	    update.set("content."+"famly", "sun");
	    mongoTemplate.updateMulti(query, update, MCol.class,"aa");
		*/
	 
	  
	}

}
