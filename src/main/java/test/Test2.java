package test;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;
import org.springframework.data.mongodb.core.MongoTemplate;

import acmr.excel.pojo.ExcelBook;

public class Test2 {
	public static void main(String[] args) {
       ExcelBook excel = new ExcelBook();
       ApplicationContext applicationContext = new ClassPathXmlApplicationContext("/config/mongodb.xml");
		//ApplicationContext app = new ClassPathXmlApplicationContext("/config/applicationContext-core.xml");
		MongoTemplate mongoTemplate = (MongoTemplate) applicationContext.getBean("mongoTemplate");
		mongoTemplate.insert(excel);
       
	}
}
