package test;

import org.springframework.beans.BeansException;
import org.springframework.context.ApplicationContext;
import org.springframework.context.ApplicationContextAware;
import org.springframework.data.mongodb.core.MongoTemplate;

public class Test implements ApplicationContextAware {

	private MongoTemplate mongoTemplate;
	
	
	public MongoTemplate getMongoTemplate() {
		return mongoTemplate;
	}


	public void setMongoTemplate(MongoTemplate mongoTemplate) {
		this.mongoTemplate = mongoTemplate;
	}


	@Override
	public void setApplicationContext(ApplicationContext applicationContext) throws BeansException {
		
		 MongoTemplate mongoTemplate = applicationContext.getBean( 
			        "mongoTemplate", MongoTemplate.class); 
			    setMongoTemplate(mongoTemplate); 
		
	}

}
