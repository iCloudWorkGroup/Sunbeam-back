package com.acmr.excel.distribute;

import java.util.Map;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import com.mysql.fabric.xmlrpc.base.Array;



public class Test {

   private static Logger log = LoggerFactory.getLogger(Test.class);
	
   public void delCol(Map transMap){
	   System.out.println("aaa");
	   transMap.put("aa", "你好");
   }
   
   public void delCell(String name,Map reviceMap){
	  
	   reviceMap.clear();
	   reviceMap.put("bb", "One 你好");
   }
   public void insertCell(){
	   System.out.println("insert cell");
   }
   public void undo(){
	   
   }

   public static void main(String[] args) throws InstantiationException, IllegalAccessException, ClassNotFoundException {
	  log.info("开始");
	  Test test = new Test();
	  Target target = new Target(test,"insertCell",0);
	  Target target1 = new Target(test,"delCol",1);
	  Target target2 = new Target(test,"delCell",1);
	  Distribute dis = new Distribute();
	  dis.add(target);
	  dis.add(target1);
	  dis.add(target2);
	  dis.exec();
	  log.info("完成");
}
}
