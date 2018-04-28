/**
 * Convert.java
 *北京华通人商用信息有限公司版权所有
 */

package com.acmr.excel.convert;

import org.apache.poi.ss.usermodel.Workbook;

import com.acmr.excel.model.complete.CompleteExcel;
/**
 * excel转换接口，子类需继承
 * @author jinhr
 *
 */

public interface Convert {
	/**
	 * 转换类型
	 * @author jinhr
	 *
	 */
	public static enum ConvertType {
		ONLINEEXCEL// 在线excel
		;
	}
	/**
	 * 讲workbook转换为CompleteExcel对象
	 * @param workbook excel文件
	 * @return CompleteExcel对象
	 */
	
	CompleteExcel doConvertExcel(Workbook workbook);

}
