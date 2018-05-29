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

}