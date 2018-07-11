package com.acmr.excel.distribute;

import java.util.Map;

public class Test {
   public void speak(Map map){
	   System.out.println("aaa");
	   map.put("aa", "你好");
   }
   
   public void say(Map map){
	   System.out.println(map.get("aa"));
	   map.clear();
	   map.put("bb", "One 你好");
   }
   
   public static void main(String[] args) throws InstantiationException, IllegalAccessException, ClassNotFoundException {
	  Test test = new Test();
	  Target target = new Target(test,"speak");
	  Target target1 = new Target(test,"say");
	  Target target2 = new Target(Class.forName(One.class.getName()).newInstance(),"say");
	  Distribute dis = new Distribute();
	  dis.add(target);
	  dis.add(target1);
	  dis.add(target2);
	  dis.exec();
}
}
