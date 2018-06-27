package test;

import java.util.ArrayList;
import java.util.List;

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
      /*		query.addCriteria(Criteria.where("_id").is("cList").and("rcList.alias").is("1"));
		Update update = new Update();
		update.set("rcList.length", 120);
		mongoTemplate.updateFirst(query, update, "6d16ae1b-7191-4c4d-b67b-a49798a22d00");
		*/
	    query.addCriteria(Criteria.where("id").is("aa"));
	    List<Object> list = new ArrayList<Object>();
	    list.add(3);
	    list.add(4);
	    Update update = new Update();
	    //update.addToSet("comments", 4);
	    update.pushAll("comments", list.toArray());
	    //update.push("comments", 5);
	    mongoTemplate.updateFirst(query, update, "aa");
	    
}
} 
