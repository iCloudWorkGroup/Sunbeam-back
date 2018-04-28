package com.acmr.excel.test;

import java.util.ArrayList;
import java.util.List;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.Test;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelRow;

import com.acmr.excel.model.ColorSet;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.service.HandleExcelService;

/**
 * 单元格操作测试类
 * 
 * @author jinhr
 *
 */
public class CellTest {
	private HandleExcelService handleExcelService;
	private ExcelBook excelBook;

	@Before
	public void before() {
		handleExcelService = new HandleExcelService();
		excelBook = TestUtil.createNewExcel();
	}

	/**
	 * 测试批量设置单元格颜色
	 */
	@Test
	public void testBatchColorSet() {
		ColorSet colorSet = new ColorSet();
		List<Coordinate> coordinates = new ArrayList<Coordinate>();
		Coordinate coordinate1 = new Coordinate();
		coordinate1.setEndRow(2);
		coordinate1.setStartRow(0);
		coordinate1.setStartCol(0);
		coordinate1.setEndCol(2);
		Coordinate coordinate2 = new Coordinate();
		coordinate2.setEndRow(5);
		coordinate2.setStartRow(3);
		coordinate2.setStartCol(0);
		coordinate2.setEndCol(2);
		coordinates.add(coordinate1);
		coordinates.add(coordinate2);
		colorSet.setCoordinates(coordinates);
		String color = "rgb(73, 68, 41)";
		colorSet.setColor(color);
		handleExcelService.batchColorSet(colorSet, excelBook);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		for (Coordinate coordinate : coordinates) {
			for (int i = coordinate.getStartRow(); i <= coordinate.getEndRow(); i++) {
				for (int j = coordinate.getStartCol(); j <= coordinate.getEndCol(); j++) {
					ExcelCell excelCell = rowList.get(i).getCells().get(j);
					ExcelColor excelColor = excelCell.getCellstyle().getFgcolor();
					Assert.assertEquals(73, excelColor.getR());
					Assert.assertEquals(68, excelColor.getG());
					Assert.assertEquals(41, excelColor.getB());
				}
			}
		}
	}
}
