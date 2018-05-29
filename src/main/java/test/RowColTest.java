package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.data.mongodb.core.query.Criteria;
import org.springframework.data.mongodb.core.query.Query;
import org.springframework.data.mongodb.core.query.Update;

public class RowColTest {
 public static void main(String[] args) {
	
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
	    Query query = new Query();
		query.addCriteria(Criteria.where("_id").is("cList").and("rcList.alias").is("1"));
		Update update = new Update();
		update.set("rcList.length", 120);
		mongoTemplate.updateFirst(query, update, "6d16ae1b-7191-4c4d-b67b-a49798a22d00");
}
}
