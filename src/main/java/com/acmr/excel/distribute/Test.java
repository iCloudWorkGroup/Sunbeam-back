package com.acmr.excel.distribute;

import java.util.Map;

import org.apache.xbean.spring.context.ClassPathXmlApplicationContext;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.context.ApplicationContext;

import com.acmr.excel.model.ColOperate;
import com.acmr.excel.service.impl.MColServiceImpl;

public class Test {

	private static Logger log = LoggerFactory.getLogger(Test.class);

	private One one = new One();

	public void delCol(Map transMap) {
		System.out.println("aaa");
		transMap.put("aa", "你好");
	}

	public void delCell(Map reviceMap) {

		reviceMap.clear();
		reviceMap.put("bb", "One 你好");
		one.say(reviceMap);
	}

	public void insertCell() {
		System.out.println("insert cell");
	}

	public void undo() {

	}

	public static void main(String[] args) throws InstantiationException,
			IllegalAccessException, ClassNotFoundException {
		log.info("开始");

		/*
		 * Target target = new Target(Test.class,"insertCell",0); Target target1
		 * = new Target(Test.class,"delCol",1); Target target2 = new
		 * Target(Test.class,"delCell",1); Distribute dis = new Distribute();
		 * dis.add(target); dis.add(target1); dis.add(target2); dis.exec();
		 * log.info("完成");
		 */
		@SuppressWarnings("resource")
		ApplicationContext applicationContext = new ClassPathXmlApplicationContext( 
				new String[]{"config/applicationContext-core.xml", 
						"config/ActiveMQ.xml","config/spring-mvc.xml"
						});
		
		MColServiceImpl mc = (MColServiceImpl) applicationContext.getBean("mcolService");
		ColOperate colOperate = new ColOperate();
		mc.showCol(colOperate, "aa", 1);
	}
}
