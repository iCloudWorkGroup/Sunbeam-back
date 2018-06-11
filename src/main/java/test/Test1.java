package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.mongo.MExcel;

public class Test1 {

	public static void main(String[] args) {
		
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		
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
		Query query = new Query();
		query.addCriteria(Criteria.where("_id").is("rList").and("rcList.alias").is("9"));
		Update update = new Update();
		update.set("rcList.$.preAlias", "8");
		mongoTemplate.updateFirst(query, update, "0c8d6409-700f-4d18-875a-d948b4c8c07c");
	}

}
