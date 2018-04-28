package com.acmr.excel.util;

import java.util.UUID;

/**
 * UUID工具类
 * 
 * @author jinhr
 *
 */

public class UUIDUtil {
	/**
	 * 生成uuid
	 * 
	 * @return uuid
	 */
	public static String getUUID() {
		return UUID.randomUUID().toString();
	}
	
//	public static void main(String[] args) {
////		TestP p = new TestP();
////		p.setS1("s1");
////		p.setS2("s2");
////		Map<String, String> map = new HashMap<String, String>();
////		String s1 = JSON.toJSONString(p);
////		System.out.println(s1);
////		map.put(testEnum.test1.toString(), s1);
////		String s = JSON.toJSONString(map, true);
////		System.out.println(s);
////		Map<String, String> newMap = (Map<String, String>) JSON.parse(s);
////		String newS = newMap.get(testEnum.test1.toString());
////		System.out.println(newS);
////		TestP testP = (TestP)JSON.parseObject(newS, TestP.class);
////		System.out.println(testP.getS2());
////		TestP p1 = new TestP();
////		p1.setS1("s1");
////		p1.setS2("s2");
////		TestP p2 = new TestP();
////		p2.setS1("s1");
////		p2.setS2("s2");
////		List<String> pList = new ArrayList<String>();
////		pList.add(JSON.toJSONString(p1));
////		pList.add(JSON.toJSONString(p2));
////		Map<String,String> pMap = new HashMap<String,String>();
////		pMap.put("1", JSONArray.toJSONString(pList));
////		String s = JSON.toJSONString(pMap);
////		System.out.println(s);
////		Map<String,String> newMap = (Map<String,String>)JSON.parse(s);
////		String newPList = newMap.get("1");
////		List<String> newPList1 = JSONArray.parseObject(newPList, List.class);
////		String n = newPList1.get(0);
////		TestP newTestP = JSONObject.parseObject(n, TestP.class);
////		System.out.println(newTestP.getS1());
//		
//		TestP p1 = new TestP();
//		p1.setS1("s1");
//		p1.setS2("s2");
//		TestP p2 = new TestP();
//		p2.setS1("s1");
//		p2.setS2("s2");
//		List<TestP> pList = new ArrayList<TestP>();
//		pList.add(p1);
//		pList.add(p2);
//		Map<String,String> pMap = new HashMap<String,String>();
//		pMap.put("1", JSONArray.toJSONString(pList));
//		String s = JSON.toJSONString(pMap);
//		System.out.println(s);
//		Map<String,String> newMap = (Map<String,String>)JSON.parse(s);
//		String newPList = newMap.get("1");
//		List newPList1 = JSONArray.parseObject(newPList, List.class);
//		JSONObject  jp = (JSONObject)newPList1.get(0);
//		TestP newTestP = JSONObject.toJavaObject(jp, TestP.class);
//		System.out.println(newTestP.getS1());
//		
//		
//		
//		
//		
//		
//		
//		
//		
//		
//		
//		
//		
//	}
}