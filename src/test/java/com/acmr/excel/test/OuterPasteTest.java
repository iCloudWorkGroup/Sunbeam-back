package com.acmr.excel.test;

import java.util.ArrayList;
import java.util.List;

import junit.framework.Assert;

import org.junit.After;
import org.junit.Before;
import org.junit.Test;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

import com.acmr.excel.model.OuterPasteData;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.PasteService;
/**
 * 外部粘贴单元测试
 * @author jinhr
 *
 */
public class OuterPasteTest {
	private Paste paste;
	private ExcelBook excelBook;
	private PasteService pasteService;

	@Before
	public void before() {
		paste = new Paste();
		excelBook = createNewExcel();
		paste.setPasteData(new ArrayList<OuterPasteData>());
		pasteService = new PasteService();
	}

	private ExcelBook createNewExcel() {
		ExcelBook excelBook = new ExcelBook();
		ExcelSheet sheet = new ExcelSheet();
		for (int i = 1; i < 27; i++) {
			ExcelColumn column = sheet.addColumn();
			column.setWidth(69);
		}
		for (int i = 1; i < 101; i++) {
			ExcelRow row = sheet.addRow();
			row.setHeight(19);
		}
		excelBook.getSheets().add(sheet);
		return excelBook;
	}
	/**
	 * 测试区域中含有空白内容时的外部粘贴
	 */
	@Test
	public void testOuterPasteWithSpace() {
		paste.setOprRow(0);
		paste.setOprCol(0);
		paste.setRowLen(2);
		paste.setColLen(2);
		OuterPasteData outerPasteData = new OuterPasteData();
		outerPasteData.setRow(0);
		outerPasteData.setCol(0);
		outerPasteData.setContent("A1");
		OuterPasteData outerPasteDat2 = new OuterPasteData();
		outerPasteDat2.setRow(1);
		outerPasteDat2.setCol(0);
		outerPasteDat2.setContent("A2");
		paste.getPasteData().add(outerPasteData);
		paste.getPasteData().add(outerPasteDat2);
		List<ExcelRow> rowList = excelBook.getSheets().get(0).getRows();
		ExcelCell cell01 = new ExcelCell();
		cell01.setText("B1");
		ExcelCell cell02 = new ExcelCell();
		cell02.setText("C1");
		rowList.get(0).getCells().set(1, cell01);
		rowList.get(0).getCells().set(2, cell02);
		ExcelCell cell12 = new ExcelCell();
		cell12.setText("B2");
		ExcelCell cell22 = new ExcelCell();
		cell22.setText("C2");
		rowList.get(1).getCells().set(1, cell12);
		rowList.get(2).getCells().set(1, cell22);
		pasteService.data(paste, excelBook, new VersionHistory(), 1);
		rowList = excelBook.getSheets().get(0).getRows();
		Assert.assertNull(rowList.get(0).getCells().get(1));
		Assert.assertEquals("C1", rowList.get(0).getCells().get(2).getText());
		Assert.assertNull(rowList.get(1).getCells().get(1));
		Assert.assertEquals("C2", rowList.get(2).getCells().get(1).getText());
		Assert.assertEquals("A1", rowList.get(0).getCells().get(0).getText());
		Assert.assertEquals("A2", rowList.get(1).getCells().get(0).getText());
	}

	@After
	public void after() {
		paste = null;
		excelBook = null;
		pasteService = null;
	}

}
