package com.acmr.excel.model;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import com.acmr.excel.util.PropertiesReaderUtil;

public class Constant {
	public static final int SUCCESS_CODE = 200;
	public static final String SUCCESS_MSG = "操作成功";
	public static final int CACHE_INVALID_CODE = 300;
	public static final String CACHE_INVALID_MSG = "缓存失效";
	public static final String MEMCACHED_EXP_TIME = PropertiesReaderUtil
			.get("memcache.expiretime");
	public static final String queueName = PropertiesReaderUtil
			.get("activemq.queueName");
	public static final String frontName = PropertiesReaderUtil
			.get("frontName");
	public static final String outPath = PropertiesReaderUtil.get("outPath");
	// public static Map<String,Object> map = new HashMap<String,Object>();
	/**
	 * 单元格高亮
	 */
	public static final String HIGHLIGHT = "highlight";

	public static List<String> accessControlAllowOriginList = new ArrayList<String>();
	
	public static Map<String,Boolean> unRecordAction = new HashMap<String,Boolean>();
	
	public static Map<String,Boolean> unInterceptAction = new HashMap<String,Boolean>();

	static {
		accessControlAllowOriginList.add("http://localhost:4711");
		accessControlAllowOriginList.add("http://192.168.2.193:8080");
		accessControlAllowOriginList.add("http://192.168.2.241:8080");
		accessControlAllowOriginList.add("http://localhost:8080");
		accessControlAllowOriginList.add("http://192.168.2.207:8080");
		accessControlAllowOriginList.add("http://192.168.2.16:8080");
		accessControlAllowOriginList.add("http://192.168.2.234:8080");
		accessControlAllowOriginList.add("http://192.168.1.241:8080");
		accessControlAllowOriginList.add("http://192.168.2.78:8080");
		accessControlAllowOriginList.add("http://192.168.2.73:8080");
		accessControlAllowOriginList.add("http://192.168.1.194:8080");
		accessControlAllowOriginList.add("http://192.168.1.180:8080");
		accessControlAllowOriginList.add("http://192.168.3.154:8080");
		
		
		unRecordAction.put("reload", true);
		unRecordAction.put("frozen", true);
		unRecordAction.put("unFrozen", true);
		unRecordAction.put("redo", true);
		unRecordAction.put("undo", true);
		unRecordAction.put("addRow", true);
		unRecordAction.put("addCol", true);
		unRecordAction.put("hideRow", true);
		unRecordAction.put("showRow", true);
		unRecordAction.put("hideCol", true);
		unRecordAction.put("showCol", true);
		unRecordAction.put("updateLock", true);
		unRecordAction.put("updateProtect", true);
		unRecordAction.put("download", true);
		unRecordAction.put("clear_queue", true);
		
		
		unInterceptAction.put("protect", true);
		unInterceptAction.put("data", true);
		unInterceptAction.put("paste", true);
		unInterceptAction.put("copy", true);
		unInterceptAction.put("cut", true);
		unInterceptAction.put("position", true);
		unInterceptAction.put("openexcel", true);
		unInterceptAction.put("welcom", true);
		unInterceptAction.put("main", true);
		unInterceptAction.put("download", true);
		unInterceptAction.put("upload", true);
		unInterceptAction.put("reopen", true);
		unInterceptAction.put("clear_queue", true);
	}

}
