package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

import com.acmr.excel.model.mongo.MRow;

public class BorderTest {
	public static void main(String[] args) {
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		Query query = new Query();
		query.addCriteria(Criteria.where("alias").is("12").and("_class")
				.is(MRow.class.getName()).and("sheetId").is("8af74f42-cf84-410f-99b3-9dbd250027d70"));
		Update update = new Update();
	
		update.set("border.top", 1);
		
		mongoTemplate.updateMulti(query, update, MRow.class, "8af74f42-cf84-410f-99b3-9dbd250027d7");
 }
}