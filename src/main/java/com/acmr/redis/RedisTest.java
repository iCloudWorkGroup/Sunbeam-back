package com.acmr.redis;

import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;

import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.HistoryCache;

public class RedisTest {
	
	private  List<History> list;
	
	private HistoryCache cache;
	
	public static void main(String[] args) {
		RedisTest  test = new RedisTest();
		test.a();
		
		
	}
	
	
	public void a(){
		ApplicationContext applicat = new ClassPathXmlApplicationContext(
				new String[] { "config/applicationContext-core.xml",
						"config/redis.xml","config/ActiveMQ.xml" });
		Redis re = (Redis) applicat.getBean("redis");
		cache = (HistoryCache) re.get("8e4d5e06-2070-4962-bc96-40881a9d7390");
		list = cache.getList();
		/*for(int i=list.size()-2;i<list.size();i++){
			list.remove(i);
		}*/
		re.set("aa", cache);
		History history = new History();
		list.add(history);
		
		
	}
}
