package com.acmr.excel.util;

import java.util.List;

import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;

import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
/**
 * 二分查找又称折半查找，它是一种效率较高的查找方法。 　　【二分查找要求】：1.必须采用顺序存储结构 2.必须按关键字大小有序排列。
 * 
 * @author Administrator
 *
 */

public class BinarySearch {

	/**
	 * 根据列别名获得列索引
	 * 
	 * @param columnList
	 * @param colAlais
	 * @return
	 */

	public static int getColIndexByColAlais(List<ExcelColumn> columnList,
			String colAlais) {
		int colIndex = -1;
		for (int i = 0; i < columnList.size(); i++) {
			if (columnList.get(i).getCode().equals(colAlais)) {
				colIndex = i;
				break;
			}
		}
		return colIndex;
	}

	/**
	 * 根据行别名获得行索引
	 * 
	 * @param rowList
	 * @param rowAlais
	 * @return
	 */

	public static int getRowIndexByRowAlais(List<ExcelRow> rowList,
			String rowAlais) {
		int rowIndex = -1;
		for (int i = 0; i < rowList.size(); i++) {
			if (rowList.get(i).getCode().equals(rowAlais)) {
				rowIndex = i;
				break;
			}
		}
		return rowIndex;
	}

	/**
	 * 二分查询行索引
	 * 
	 * @param glyList
	 * @param top
	 * @return
	 */

	public static int rowsBinarySearch(List<Gly> glyList, int top) {
		int index = 0; // 检索的时候
		int start = 0; // 用start和end两个索引控制它的查询范围
		int end = glyList.size() - 1;
		int result = -1;
		// count = 0;
		for (int i = 0; i < glyList.size(); i++) {
			// count++;
			index = (start + end) / 2;
			if (glyList.size() - 1 == i) {
				// System.out.println(num + "找到了，在数组下标为" + end + "的地方,查找了" +
				// count + "次。");
				result = end;
				break;
			} else if (glyList.get(index).getTop() < top) {
				if (glyList.get(index + 1).getTop() > top) {
					// System.out.println(array[index] + "找到了，在数组下标为" + index +
					// "的地方,查找了" + count + "次。");
					result = index;
					break;
				}
				start = index;
			} else if (glyList.get(index).getTop() > top) {
				end = index;
			} else {
				// System.out.println(array[index] + "找到了，在数组下标为" + index +
				// "的地方,查找了" + count + "次。");
				result = index;
				break;
			}
		}
		return result;
	}
	/**
	 * 二分查询列索引
	 * 
	 * @param glyList
	 * @param top
	 * @return
	 */

	public static int colsBinarySearch(List<Glx> glxList, int left) {
		int index = 0; // 检索的时候
		int start = 0; // 用start和end两个索引控制它的查询范围
		int end = glxList.size() - 1;
		int result = -1;
		// count = 0;
		for (int i = 0; i < glxList.size(); i++) {
			// count++;
			index = (start + end) / 2;
			if (glxList.size() - 1 == i) {
				// System.out.println(num + "找到了，在数组下标为" + end + "的地方,查找了" +
				// count + "次。");
				result = end;
				break;
			} else if (glxList.get(index).getLeft() < left) {
				if (glxList.get(index + 1).getLeft() > left) {
					// System.out.println(array[index] + "找到了，在数组下标为" + index +
					// "的地方,查找了" + count + "次。");
					result = index;
					break;
				}
				start = index;
			} else if (glxList.get(index).getLeft() > left) {
				end = index;
			} else {
				// System.out.println(array[index] + "找到了，在数组下标为" + index +
				// "的地方,查找了" + count + "次。");
				result = index;
				break;
			}
		}
		return result;
	}
}