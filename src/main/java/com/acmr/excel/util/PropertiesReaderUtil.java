package com.acmr.excel.util;

import java.io.FileOutputStream;
import java.io.InputStream;
import java.util.Properties;

import acmr.data.DataQueryPool;

/**
 * 获取系统param.properties参数文件内容，并存入内存，实现参数全局使用
 * 
 *
 */
public class PropertiesReaderUtil {

	private static Properties p = null;

	public PropertiesReaderUtil() {

	}

	/**
	 * 获取acmr.properties文件内容
	 * 
	 * @param key
	 * @return
	 */
	public static String get(String key) {
		if (p == null) {
			InputStream inputStream = DataQueryPool.class.getClassLoader().getResourceAsStream("properties/param.properties");
			p = new Properties();
			try {
				p.load(inputStream);
				inputStream.close();
			} catch (Exception e) {
				//System.out.println("init PropertiesReader error!" + e);
				return "";
			}
		}
		return p.getProperty(key);
	}

	/**
	 * 返回　Properties
	 * 
	 * @param fileName
	 *            文件名　(注意：加载的是src下的文件,如果在某个包下．请把包名加上)
	 * @param
	 * @return
	 */
	public static Properties getProperties() {
		return p;
	}

	/**
	 * 写入properties信息
	 * 
	 * @param key
	 *            名称
	 * @param value
	 *            值
	 */
	public static void modifyProperties(String key, String value) {
		try {
			// 从输入流中读取属性列表（键和元素对）
			Properties prop = getProperties();
			prop.setProperty(key, value);
			String path = PropertiesReaderUtil.class.getResource(
					"config/acmr.properties").getPath();
			FileOutputStream outputFile = new FileOutputStream(path);
			prop.store(outputFile, "modify");
			outputFile.close();
			outputFile.flush();
		} catch (Exception e) {
		}
	}
}
