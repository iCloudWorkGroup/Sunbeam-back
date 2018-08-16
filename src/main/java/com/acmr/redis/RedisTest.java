package com.acmr.redis;

import java.util.ArrayList;
import java.util.List;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.springframework.context.ApplicationContext;

public class RedisTest {
	public static void main(String[] args) {
		ApplicationContext applicat = new ClassPathXmlApplicationContext(
				new String[] { "config/applicationContext-core.xml",
						"config/redis.xml","config/ActiveMQ.xml" });
		Redis re = (Redis) applicat.getBean("redis");
		List<String> list = new ArrayList<String>();
		list.add("a");
		re.set("list", list);
		list.add("1111");
		System.out.println(((List)re.get("list")).get(0));
		System.out.println(((List)re.get("list")).get(1));
		re.set("bb", 123);
		System.out.println(re.get("bb"));
		re.del("bb");
		System.out.println(re.get("bb"));
		re.set("aa","你好");
		System.out.println(re.get("aa"));
		
		
	}
}
