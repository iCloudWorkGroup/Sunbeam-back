package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MColDao;
import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.mongo.MBook;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MRowColList;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MBookService;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.CellFormateUtil;
import com.acmr.excel.util.ExcelUtil;

import acmr.excel.ExcelException;
import acmr.excel.ExcelHelper;
import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.util.ListHashMap;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.excel.pojo.Excelborder;

@Service("mbookService")
public class MBookServiceImpl implements MBookService {

	@Resource
	private BaseDao baseDao;

	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MRowDao mrowDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MRowColCellDao mrowColCellDao;
	@Resource
	private MSheetDao msheetDao;

	@Override
	public boolean saveExcelBook(ExcelBook book, String excelId) {
		if (book == null) {
			return false;
		}
		MBook mbook = new MBook();
		mbook.setId(excelId);
		mbook.setName("新建Excel");
		baseDao.insert(excelId, mbook);// 存储MBook对象
		List<ExcelSheet> excelSheets = book.getSheets();
		ExcelSheet excelSheet = excelSheets.get(0);
		MSheet msheet = new MSheet();
		String sheetId = excelId + 0;
		msheet.setId(sheetId);
		msheet.setStep(0);
		msheet.setMaxrow(excelSheet.getMaxrow());
		msheet.setMaxcol(excelSheet.getMaxcol());
		msheet.setSheetName(excelSheet.getName());
		ExcelSheetFreeze freeze = excelSheet.getFreeze();
		if (null != freeze) {
			msheet.setFreeze(freeze.isFreeze());
			msheet.setViewRowAlias(Integer
					.toString(freeze.getFirstrow() + 1 - freeze.getRow()));
			msheet.setViewColAlias(Integer
					.toString(freeze.getFirstcol() + 1 - freeze.getCol()));
			msheet.setRowAlias(Integer.toString(freeze.getFirstrow() + 1));
			msheet.setColAlias(Integer.toString(freeze.getFirstcol() + 1));
		} else {
			msheet.setFreeze(false);
			msheet.setViewColAlias("1");
			msheet.setViewRowAlias("1");
		}
		baseDao.insert(excelId, msheet);// 存储MSheet对象

		List<ExcelColumn> cols = excelSheet.getCols();// 得到所有的列
		List<MCol> mcols = new ArrayList<MCol>();// 存贮列样式
		MRowColList cList = new MRowColList();// 用于存贮简化的列
		cList.setId("cList");
		cList.setSheetId(sheetId);
		getMCol(cols, mcols, cList, sheetId);
		// mongoTemplate.insert(mcols, excelId);
		baseDao.insertList(excelId, mcols);// 存列表样式
		baseDao.insert(excelId, cList);// 向数据库存贮简化的列信息

		List<ExcelRow> rows = excelSheet.getRows();
		List<Object> tempList = new ArrayList<Object>();// 存储mcell对象及关系表对象
		List<MRow> mrows = new ArrayList<MRow>();
		MRowColList rList = new MRowColList();
		rList.setId("rList");
		rList.setSheetId(sheetId);
		getMRow(rows, mrows, rList, tempList, sheetId);
		baseDao.insert(excelId, rList);// 向数据库存贮行信息
		// mongoTemplate.insert(mrows,excelId);
		baseDao.insertList(excelId, mrows);// 存储行样式

		List<Sheet> sheets = book.getNativeSheet();// 找出合并的单元格
		getMergeCell(sheets, tempList, sheetId);

		long start = System.currentTimeMillis();
		// mongoTemplate.insert(tempList, excelId);
		baseDao.insertList(excelId, tempList);
		long end = System.currentTimeMillis() - start;
		System.out.println("存储时间为:" + end);

		return true;

	}

	private void getMCol(List<ExcelColumn> cols, List<MCol> mcols,
			MRowColList cList, String sheetId) {

		for (int i = 0; i < cols.size(); i++) {
			ExcelColumn excelColumn = cols.get(i);
			ExcelCellStyle style = excelColumn.getCellstyle();
			ExcelFont font = style.getFont();
			MCol mcol = new MCol();
			RowCol rc = new RowCol();

			rc.setAlias(excelColumn.getCode());
			if (i > 0) {

				rc.setPreAlias(cols.get(i - 1).getCode());// 存贮它前一行的别名
			}
			mcol.setSheetId(sheetId);
			mcol.setAlias(excelColumn.getCode());
			boolean hidden = excelColumn.isColumnhidden();
			if (hidden) {
				mcol.setHidden(true);
				mcol.setWidth(excelColumn.getWidth());
				rc.setLength(0);
			} else {
				mcol.setWidth(excelColumn.getWidth());
				mcol.setHidden(false);
				rc.setLength(excelColumn.getWidth());
			}

			cList.getRcList().add(rc);// 存简化列属性

			OperProp operProp = mcol.getProps();
			mcols.add(mcol);

			Content content = operProp.getContent();
			switch (style.getAlign()) {
			case 1:
				content.setAlignRow(null);
				break;
			case 2:
				content.setAlignRow("center");
				break;
			case 3:
				content.setAlignRow("right");
				break;

			default:

				break;
			}
			switch (style.getValign()) {
			case 0:
				content.setAlignCol(null);
				break;
			case 1:
				content.setAlignCol("middle");
				break;
			case 2:
				content.setAlignCol("bottom");
				break;

			default:

				break;
			}

			if (700 == font.getBoldweight()) {
				content.setWeight(true);
			} else {
				content.setWeight(null);
			}
			if (null == font.getColor()) {
				content.setColor(null);
			} else {
				content.setColor(
						CellFormateUtil.getBackground(font.getColor()));
			}

			switch (font.getFontname()) {
			case "微软雅黑":
				content.setFamily("microsoft Yahei");
				break;
			case "宋体":
				content.setFamily(null);
				break;
			case "SimSun":
				content.setFamily(null);
				break;
			case "黑体":
				content.setFamily("SimHei");
				break;
			case "楷体":
				content.setFamily("KaiTi");
				break;
			default:
				content.setFamily(font.getFontname());
				break;
			}
			content.setDisplayTexts(null);
			if (font.isItalic()) {
				content.setItalic(true);
			} else {
				content.setItalic(null);
			}
			if (font.getSize() == 220) {
				content.setSize(null);
			} else {
				content.setSize(font.getSize() / 20 + "");
			}
			content.setTexts(null);
			if (0 == (int) font.getUnderline()) {
				content.setUnderline(null);
			} else {
				content.setUnderline((int) font.getUnderline());
			}
			if (style.isWraptext()) {
				content.setWordWrap(style.isWraptext());
			} else {
				content.setWordWrap(null);
			}

			content.setType(null);
			content.setExpress(style.getDataformat());

			if (null == style.getFgcolor()) {
				content.setBackground(null);
			} else {
				content.setBackground(
						CellFormateUtil.getBackground(style.getFgcolor()));
			}
			if (style.isLocked()) {
				content.setLocked(null);
			} else {
				content.setLocked(style.isLocked());

			}

			CustomProp customProp = operProp.getCustomProp();
			customProp.setComment(null);

			Border border = operProp.getBorder();
			Excelborder excelTopborder = style.getTopborder();
			if (excelTopborder != null) {
				short topBorder = excelTopborder.getSort();
				if (topBorder == 2) {
					border.setTop(2);
				} else if (topBorder == 1) {
					border.setTop(1);
				} else {
					border.setTop(null);
				}
			}
			Excelborder excelBottomborder = style.getBottomborder();
			if (excelBottomborder != null) {
				short bottomBorder = excelBottomborder.getSort();
				if (bottomBorder == 2) {
					border.setBottom(2);
				} else if (bottomBorder == 1) {
					border.setBottom(1);
				} else {
					border.setBottom(null);
				}
			}
			Excelborder excelLeftborder = style.getLeftborder();
			if (excelLeftborder != null) {
				short leftBorder = excelLeftborder.getSort();
				if (leftBorder == 2) {
					border.setLeft(2);
				} else if (leftBorder == 1) {
					border.setLeft(1);
				} else {
					border.setLeft(null);
				}
			}
			Excelborder excelRightborder = style.getRightborder();
			if (excelRightborder != null) {
				short rightBorder = excelRightborder.getSort();
				if (rightBorder == 2) {
					border.setRight(2);
				} else if (rightBorder == 1) {
					border.setRight(1);
				} else {
					border.setRight(null);
				}
			}

		}
	}

	private void getMRow(List<ExcelRow> rows, List<MRow> mrows,
			MRowColList rList, List<Object> tempList, String sheetId) {
		for (int i = 0; i < rows.size(); i++) {
			ExcelRow excelRow = rows.get(i);
			if (excelRow != null) {

				ExcelCellStyle style = excelRow.getCellstyle();
				ExcelFont font = style.getFont();
				MRow mrow = new MRow();
				RowCol rc = new RowCol();

				rc.setAlias(excelRow.getCode());
				if (i > 0) {

					rc.setPreAlias(rows.get(i - 1).getCode());// 存贮它前一行的别名
				}

				mrow.setSheetId(sheetId);
				mrow.setAlias(excelRow.getCode());
				boolean hidden = excelRow.isRowhidden();
				if (hidden) {
					mrow.setHidden(true);
					mrow.setHeight(excelRow.getHeight());
					rc.setLength(0);
				} else {
					mrow.setHeight(excelRow.getHeight());
					mrow.setHidden(false);
					rc.setLength(excelRow.getHeight());
				}

				rList.getRcList().add(rc);// 存简化行属性

				OperProp prop = mrow.getProps();
				mrows.add(mrow);

				Content content = prop.getContent();

				switch (style.getAlign()) {
				case 1:
					content.setAlignRow(null);
					break;
				case 2:
					content.setAlignRow("center");
					break;
				case 3:
					content.setAlignRow("right");
					break;

				default:

					break;
				}
				switch (style.getValign()) {
				case 0:
					content.setAlignCol(null);
					break;
				case 1:
					content.setAlignCol("middle");
					break;
				case 2:
					content.setAlignCol("bottom");
					break;

				default:

					break;
				}

				if (700 == font.getBoldweight()) {
					content.setWeight(true);
				} else {
					content.setWeight(null);
				}
				if (null == font.getColor()) {
					content.setColor(null);
				} else {
					content.setColor(
							CellFormateUtil.getBackground(font.getColor()));
				}

				switch (font.getFontname()) {
				case "微软雅黑":
					content.setFamily("microsoft Yahei");
					break;
				case "宋体":
					content.setFamily(null);
					break;
				case "SimSun":
					content.setFamily(null);
					break;
				case "黑体":
					content.setFamily("SimHei");
					break;
				case "楷体":
					content.setFamily("KaiTi");
					break;
				default:
					content.setFamily(font.getFontname());
					break;
				}
				content.setDisplayTexts(null);
				if (font.isItalic()) {
					content.setItalic(true);
				} else {
					content.setItalic(null);
				}
				if (font.getSize() == 220) {
					content.setSize(null);
				} else {
					content.setSize(font.getSize() / 20 + "");
				}
				content.setTexts(null);
				if (0 == (int) font.getUnderline()) {
					content.setUnderline(null);
				} else {
					content.setUnderline((int) font.getUnderline());
				}
				if (style.isWraptext()) {
					content.setWordWrap(style.isWraptext());
				} else {
					content.setWordWrap(null);
				}
				content.setType(null);
				content.setExpress(style.getDataformat());
				if (null == style.getFgcolor()) {
					content.setBackground(null);
				} else {
					content.setBackground(
							CellFormateUtil.getBackground(style.getFgcolor()));
				}
				if (style.isLocked()) {
					content.setLocked(null);
				} else {

					content.setLocked(style.isLocked());
				}

				CustomProp customProp = prop.getCustomProp();
				customProp.setComment(null);

				Border border = prop.getBorder();
				Excelborder excelTopborder = style.getTopborder();
				if (excelTopborder != null) {
					short topBorder = excelTopborder.getSort();
					if (topBorder == 2) {
						border.setTop(2);
					} else if (topBorder == 1) {
						border.setTop(1);
					} else {
						border.setTop(null);
					}
				}
				Excelborder excelBottomborder = style.getBottomborder();
				if (excelBottomborder != null) {
					short bottomBorder = excelBottomborder.getSort();
					if (bottomBorder == 2) {
						border.setBottom(2);
					} else if (bottomBorder == 1) {
						border.setBottom(1);
					} else {
						border.setBottom(null);
					}
				}
				Excelborder excelLeftborder = style.getLeftborder();
				if (excelLeftborder != null) {
					short leftBorder = excelLeftborder.getSort();
					if (leftBorder == 2) {
						border.setLeft(2);
					} else if (leftBorder == 1) {
						border.setLeft(1);
					} else {
						border.setLeft(null);
					}
				}
				Excelborder excelRightborder = style.getRightborder();
				if (excelRightborder != null) {
					short rightBorder = excelRightborder.getSort();
					if (rightBorder == 2) {
						border.setRight(2);
					} else if (rightBorder == 1) {
						border.setRight(1);
					} else {
						border.setRight(null);
					}
				}

				// 存储没有合并的单元格属性
				List<ExcelCell> cells = excelRow.getCells();
				for (int j = 0; j < cells.size(); j++) {
					ExcelCell cell = cells.get(j);
					if (cell != null && cell.getColspan() < 2
							&& cell.getRowspan() < 2) {// 判断是否合并单元格
						int row = i + 1;
						int col = j + 1;
						MCell mcell = getMCell(cell, sheetId, row, col);
						tempList.add(mcell);
						// 关系映射表
						MRowColCell mrcl = new MRowColCell();
						mrcl.setSheetId(sheetId);
						mrcl.setRow(row + "");
						mrcl.setCol(col + "");
						mrcl.setCellId(row + "_" + col);
						tempList.add(mrcl);
					}
				}

			}
		}

	}

	private MCell getMCell(ExcelCell excelCell, String sheetId, int row,
			int col) {
		ExcelCellStyle style = excelCell.getCellstyle();
		ExcelFont font = style.getFont();
		MCell mcell = new MCell();
		mcell.setSheetId(sheetId);
		mcell.setId(row + "_" + col);
		mcell.setRowspan(excelCell.getRowspan());
		mcell.setColspan(excelCell.getColspan());

		Border border = mcell.getBorder();
		Excelborder excelTopborder = style.getTopborder();
		if (excelTopborder != null) {
			short topBorder = excelTopborder.getSort();
			if (topBorder == 2) {
				border.setTop(2);
			} else if (topBorder == 1) {
				border.setTop(1);
			} else {
				border.setTop(0);
			}
		}

		Excelborder excelBottomborder = style.getBottomborder();
		if (excelBottomborder != null) {
			short bottomBorder = excelBottomborder.getSort();
			if (bottomBorder == 2) {
				border.setBottom(2);
			} else if (bottomBorder == 1) {
				border.setBottom(1);
			} else {
				border.setBottom(0);
			}
		}
		Excelborder excelLeftborder = style.getLeftborder();
		if (excelLeftborder != null) {
			short leftBorder = excelLeftborder.getSort();
			if (leftBorder == 2) {
				border.setLeft(2);
			} else if (leftBorder == 1) {
				border.setLeft(1);
			} else {
				border.setLeft(0);
			}
		}
		Excelborder excelRightborder = style.getRightborder();
		if (excelRightborder != null) {
			short rightBorder = excelRightborder.getSort();
			if (rightBorder == 2) {
				border.setRight(2);
			} else if (rightBorder == 1) {
				border.setRight(1);
			} else {
				border.setRight(0);
			}
		}

		Content content = mcell.getContent();
		switch (style.getAlign()) {
		case 1:
			content.setAlignRow("left");
			break;
		case 2:
			content.setAlignRow("center");
			break;
		case 3:
			content.setAlignRow("right");
			break;

		default:

			break;
		}
		switch (style.getValign()) {
		case 0:
			content.setAlignCol("top");
			break;
		case 1:
			content.setAlignCol("middle");
			break;
		case 2:
			content.setAlignCol("bottom");
			break;

		default:

			break;
		}

		if (700 == font.getBoldweight()) {
			content.setWeight(true);
		} else {
			content.setWeight(false);
		}

		ExcelColor fontColor = font.getColor();
		if (fontColor != null) {
			int[] rgb = fontColor.getRGBInt();
			content.setColor(ExcelUtil.getRGB(rgb));

		}

		switch (font.getFontname()) {
		case "微软雅黑":
			content.setFamily("microsoft Yahei");
			break;
		case "宋体":
			content.setFamily("SimSun");
			break;
		case "黑体":
			content.setFamily("SimHei");
			break;
		case "楷体":
			content.setFamily("KaiTi");
			break;
		default:
			content.setFamily(font.getFontname());
			break;
		}

		content.setItalic(font.isItalic());
		content.setSize(font.getSize() / 20 + "");
		content.setTexts(excelCell.getText());
		CellFormateUtil.setShowText(excelCell, content);

		content.setUnderline((int) font.getUnderline());
		content.setWordWrap(style.isWraptext());
		content.setType(CellFormateUtil.TypeToMtype(excelCell.getType()));
		content.setExpress(style.getDataformat());
		ExcelColor color = style.getFgcolor();
		if (color != null) {
			int[] rgb = color.getRGBInt();
			content.setBackground(ExcelUtil.getRGB(rgb));
		} else {
			content.setBackground(null);
		}

		content.setLocked(style.isLocked());

		CustomProp customProp = mcell.getCustomProp();
		customProp.setComment(excelCell.getMemo());
		String highLight = excelCell.getExps().get(Constant.HIGHLIGHT);
		if ("true".equals(highLight)) {
			customProp.setHighlight(true);
		} else {
			customProp.setHighlight(false);
		}

		return mcell;
	}

	/***
	 * 获取合并单元格MCell对象及其映射关系MRowColCell
	 * 
	 * @param sheets
	 * @param tempList
	 * @param relation
	 */
	private void getMergeCell(List<Sheet> sheets, List<Object> tempList,
			String sheetId) {

		for (int i = 0; i < sheets.size(); i++) {
			Sheet sh = sheets.get(i);
			int mergeCount = sh.getNumMergedRegions();
			for (int j = 0; j < mergeCount; j++) {

				CellRangeAddress ca = sh.getMergedRegion(j);
				int firstRow = ca.getFirstRow();
				int firstCol = ca.getFirstColumn();
				int lastRow = ca.getLastRow();
				int lastCol = ca.getLastColumn();
				Row row = sh.getRow(firstRow);
				XSSFCell cell = (XSSFCell) row.getCell(firstCol);
				ExcelSheet es = new ExcelSheet();
				ExcelCell excell = es.getExcelCell(cell);
				// 合并单元格的最后一个cell，将右、下边框的属性赋值给合并单元格
				Row lrow = sh.getRow(lastRow);
				XSSFCell lastcell = (XSSFCell) lrow.getCell(lastCol);
				XSSFCellStyle cs = lastcell.getCellStyle();
				excell.getCellstyle().setRightborder(ExcelHelper.getExcelBorder(
						cs.getBorderRight(), cs.getRightBorderXSSFColor()));
				excell.getCellstyle().setBottomborder(
						ExcelHelper.getExcelBorder(cs.getBorderBottom(),
								cs.getBottomBorderXSSFColor()));

				int rid = firstRow + 1;
				int cid = firstCol + 1;
				excell.setRowspan(lastRow - firstRow + 1);
				excell.setColspan(lastCol - firstCol + 1);
				MCell mcell = getMCell(excell, sheetId, rid, cid);
				tempList.add(mcell);
				// 关系映射表
				for (int k = firstRow; k < lastRow + 1; k++) {
					for (int l = firstCol; l < lastCol + 1; l++) {
						// 多个单元格指向一个单元对象
						MRowColCell mrcl = new MRowColCell();
						mrcl.setSheetId(sheetId);
						mrcl.setRow((k + 1) + "");
						mrcl.setCol((l + 1) + "");
						mrcl.setCellId(rid + "_" + cid);
						tempList.add(mrcl);
					}
				}

			}
		}

	}

	@Override
	public CompleteExcel reload(String excelId, String sheetId,
			Integer rowBegin, Integer rowEnd, Integer colBegin, Integer colEnd,
			int type) {

		// 将步骤重置为0
		if (type == 0) {
			msheetDao.updateStep(excelId, sheetId, 0);
		}
		CompleteExcel excel = new CompleteExcel();
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		if (type == 0) {
			addRowOrCol(excelId, sheetId, sortRList, sortCList, rowEnd, colEnd);
		}

		Map<String, Integer> rMap = new HashMap<String, Integer>();
		for (int i = 0; i < sortRList.size(); i++) {
			rMap.put(sortRList.get(i).getAlias(), i);
		}
		Map<String, Integer> cMap = new HashMap<String, Integer>();
		for (int i = 0; i < sortCList.size(); i++) {
			cMap.put(sortCList.get(i).getAlias(), i);
		}
		int rowBeginIndex = getIndexByPixelBegin(sortRList, rowBegin);
		int rowEndIndex = getIndexByPixel(sortRList, rowEnd);
		int colBeginIndex = getIndexByPixelBegin(sortCList, colBegin);
		int colEndIndex = getIndexByPixel(sortCList, colEnd);
		if(rowBeginIndex == -1||colBeginIndex == -1){
			return excel;
		}
		int top;
		int left;
		if (rowBeginIndex == 0) {
			top = 0;
		} else {
			RowCol rc = sortRList.get(rowBeginIndex);
			top = rc.getTop();
		}
		if (colBeginIndex == 0) {
			left = 0;
		} else {
			RowCol rc = sortCList.get(colBeginIndex);
			left = rc.getTop();
		}
		List<String> rowList = new ArrayList<String>();

		List<String> colList = new ArrayList<String>();

		if (sortRList.size() > 0) {
			for (int i = rowBeginIndex; i < rowEndIndex + 1; i++) {
				rowList.add(sortRList.get(i).getAlias());
			}
		}
		if (sortCList.size() > 0) {
			for (int j = colBeginIndex; j < colEndIndex + 1; j++) {
				colList.add(sortCList.get(j).getAlias());
			}
		}

		// 行样式
		List<MRow> rList = mrowDao.getMRowList(excelId, sheetId, rowList);
		Map<String, MRow> rowMap = new HashMap<String, MRow>();
		for (MRow mr : rList) {
			rowMap.put(mr.getAlias(), mr);
		}
		List<Gly> yList = new ArrayList<Gly>();
		for (int i = 0; i < rowList.size(); i++) {
			String alias = rowList.get(i);
			Gly gly = new Gly(rowMap.get(alias));
			gly.setSort(i + rowBeginIndex);
			gly.setTop(getRowTop(yList, i, top));
			yList.add(gly);
		}

		// 列样式
		List<MCol> cList = mcolDao.getMColList(excelId, sheetId, colList);
		Map<String, MCol> colMap = new HashMap<String, MCol>();
		for (MCol mc : cList) {
			colMap.put(mc.getAlias(), mc);
		}
		List<Glx> xList = new ArrayList<Glx>();
		for (int i = 0; i < colList.size(); i++) {
			String alias = colList.get(i);
			Glx glx = new Glx(colMap.get(alias));
			glx.setSort(i + colBeginIndex);
			glx.setLeft(getColLeft(xList, i, left));
			xList.add(glx);
		}

		// 关系表
		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> cellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		List<OneCell> ceList = new ArrayList<OneCell>();
		for (MCell mc : cellList) {
			OneCell ce = new OneCell(mc);
			String[] ids = mc.getId().split("_");
			Occupy oc = ce.getOccupy();
			if (mc.getRowspan() == 1 && mc.getColspan() == 1) {
				oc.getRow().add(ids[0]);
				oc.getCol().add(ids[1]);
			} else {
				int indexRow = rMap.get(ids[0]);
				for (int i = 0; i < mc.getRowspan(); i++) {

					oc.getRow().add(sortRList.get(indexRow).getAlias());
					indexRow++;
				}
				int indexCol = cMap.get(ids[1]);
				for (int i = 0; i < mc.getColspan(); i++) {

					oc.getCol().add(sortCList.get(indexCol).getAlias());
					indexCol++;
				}

			}
			ceList.add(ce);
		}

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		SheetElement sheet = new SheetElement(msheet);
		if (sortRList.size() > 0) {
			RowCol row = sortRList.get(sortRList.size() - 1);
			sheet.setMaxRowPixel(row.getTop() + row.getLength() + 1);
		}
		if (sortCList.size() > 0) {
			RowCol col = sortCList.get(sortCList.size() - 1);
			sheet.setMaxColPixel(col.getTop() + col.getLength() + 1);
		}

		sheet.setCells(ceList);
		sheet.setGridLineRow(yList);
		sheet.setGridLineCol(xList);
		excel.getSheets().add(sheet);
		excel.setName("new Book");

		return excel;

	}

	@Override
	public ExcelBook reloadExcelBook(String excelId, Integer step) {
		String sheetId = excelId + 0;
		while (true) {
			MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
			int memStep = msheet.getStep();
			if (memStep == step) {
				
				ExcelBook book = getExcelBook(excelId,sheetId);

				return book;
			} else {
				try {
					Thread.sleep(100);
				} catch (InterruptedException e) {
					e.printStackTrace();
				}
				continue;
			}
		}
	}

	public ExcelBook getExcelBook(String excelId, String sheetId) {

		ExcelBook book = new ExcelBook();
		ListHashMap<ExcelSheet> sheets = book.getSheets();
		ExcelSheet sheet = new ExcelSheet();
		sheets.add(sheet);
		ListHashMap<ExcelColumn> cols = (ListHashMap<ExcelColumn>) sheet.getCols();
		ListHashMap<ExcelRow> rows = (ListHashMap<ExcelRow>) sheet.getRows();
		
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);

		List<String> rowList = new ArrayList<String>();
		List<String> colList = new ArrayList<String>();

		Map<String, Integer> rMap = new HashMap<String, Integer>();
		for (int i = 0; i < sortRList.size(); i++) {
			rMap.put(sortRList.get(i).getAlias(), i);
			rowList.add(sortRList.get(i).getAlias());
		}
		Map<String, Integer> cMap = new HashMap<String, Integer>();
		for (int i = 0; i < sortCList.size(); i++) {
			cMap.put(sortCList.get(i).getAlias(), i);
			colList.add(sortCList.get(i).getAlias());
		}

		// 行样式
		List<MRow> rList = mrowDao.getMRowList(excelId, sheetId, rowList);
		Map<String, MRow> rowMap = new HashMap<String, MRow>();
		for (MRow mr : rList) {
			rowMap.put(mr.getAlias(), mr);
		}

		// 列样式
		List<MCol> cList = mcolDao.getMColList(excelId, sheetId, colList);
		Map<String, MCol> colMap = new HashMap<String, MCol>();
		for (MCol mc : cList) {
			colMap.put(mc.getAlias(), mc);
		}

		// 关系表
		List<MRowColCell> relationList = mrowColCellDao
				.getMRowColCellList(excelId, sheetId, rowList, colList);
		List<String> cellIdList = new ArrayList<String>();
		Map<String,String> relationMap = new HashMap<String,String>();
		for (MRowColCell mrcc : relationList) {
			cellIdList.add(mrcc.getCellId());
			relationMap.put(mrcc.getRow()+"_"+mrcc.getCol(), mrcc.getCellId());
		}
		List<MCell> cellList = mcellDao.getMCellList(excelId, sheetId,
				cellIdList);
		Map<String, MCell> cellMap = new HashMap<String, MCell>();
		for(MCell mc:cellList){
			cellMap.put(mc.getId(), mc);
		}
		
		for(int i=0;i<rowList.size();i++){
			String rowId = rowList.get(i);
			MRow mrow = rowMap.get(rowId);
			ExcelRow row = getExcelRow(mrow);
			rows.add(row);
			
			List<ExcelCell> cells = row.getCells();
			
			
			for(int j=0;j<colList.size();j++){
				String colId = colList.get(j);
				if(i == 0){
					MCol mcol = colMap.get(colId);
					ExcelColumn col = getExcelCol(mcol);
					cols.add(col);
				}
				
				String id = rowId+"_"+colId;
				String cellId = relationMap.get(id);
				if(null == cellId){
					cells.add(null);
				}else{
					MCell mc = cellMap.get(cellId);
					ExcelCell cell = getExcelCell(mc);
					cells.add(cell);
				}
				
			}
		}

		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		if(null == msheet.getSheetName()){
			sheet.setName("Sheet1");
		}else{
			sheet.setName(msheet.getSheetName());
		}
		ExcelSheetFreeze freeze = new ExcelSheetFreeze();
		if (null != msheet.getFreeze()&&msheet.getFreeze()) {
			freeze.setFreeze(msheet.getFreeze());
			int firstRow = rMap.get(msheet.getRowAlias());
			freeze.setFirstrow(firstRow);
			int firstCol = cMap.get(msheet.getColAlias());
			freeze.setFirstcol(firstCol);
			int viewRow = rMap.get(msheet.getViewRowAlias());
			int viewCol = cMap.get(msheet.getViewColAlias());
			freeze.setRow(firstRow-viewRow);
			freeze.setCol(firstCol-viewCol);
			
		} else {
		    freeze.setFreeze(false);
		}
		sheet.setFreeze(freeze);
		sheet.setMaxcol(colList.size());
		sheet.setMaxrow(rowList.size());
		if(null != msheet.getProtect()){
			sheet.setProtect(msheet.getProtect());
			sheet.setPassword(msheet.getPasswd());
		}else{
			sheet.setProtect(false);
		}

		return book;

	}

	private ExcelCell getExcelCell(MCell mcell) {
		ExcelCell cell = new ExcelCell();

		Content content = mcell.getContent();
		Border border = mcell.getBorder();
		CustomProp cp = mcell.getCustomProp();

		cell.setRowspan(mcell.getRowspan());
		cell.setColspan(mcell.getColspan());
		cell.setText(content.getDisplayTexts());
		cell.setValue(content.getTexts());
		if (null == content.getType()) {
			cell.setType(CELLTYPE.STRING);
		} else {
			cell.setType(CellFormateUtil.MTypeToType(content.getType()));
		}

		cell.setMemo(cp.getComment());

		ExcelCellStyle style = new ExcelCellStyle();
		if(null == content.getAlignRow()){
			style.setAlign((short) 1);
		}else {
			switch (content.getAlignRow()) {
			case "left":
				style.setAlign((short) 1);
				break;
			case "center":
				style.setAlign((short) 2);
				break;
			case "right":
				style.setAlign((short) 3);
				break;
			default:
				style.setAlign((short) 2);
				break;
			}
		}
		
		if(null == content.getAlignCol()){
			style.setValign((short) 0);
		}else{
			switch (content.getAlignCol()) {
			case "top":
				style.setValign((short) 0);
				break;
			case "middle":
				style.setValign((short) 1);
				break;
			case "bottom":
				style.setValign((short) 2);
				break;
			default:
				style.setValign((short) 1);
				break;
			}
		}
		Excelborder topborder = new Excelborder();
		// ExcelColor topColor = ExcelUtil.getColor(border.getTop());
		if (null == border.getTop()) {
			topborder.setSort((short) 0);
		} else {
			topborder.setSort(border.getTop().shortValue());
		}
		style.setTopborder(topborder);
		Excelborder leftborder = new Excelborder();
		if (null == border.getLeft()) {
			leftborder.setSort((short) 0);
		} else {
			leftborder.setSort(border.getLeft().shortValue());
		}
		style.setLeftborder(leftborder);
		Excelborder rightborder = new Excelborder();
		if (null == border.getRight()) {
			rightborder.setSort((short) 0);
		} else {
			rightborder.setSort(border.getRight().shortValue());
		}
		style.setRightborder(rightborder);
		Excelborder bottomborder = new Excelborder();
		if (null == border.getBottom()) {
			bottomborder.setSort((short) 0);
		} else {
			bottomborder.setSort(border.getBottom().shortValue());
		}
		style.setBottomborder(bottomborder);

		ExcelFont font = new ExcelFont();
		if(null == content.getFamily()){
			font.setFontname("宋体");
		}else{
			
			switch (content.getFamily()) {
			case "microsoft Yahei":
				
				font.setFontname("微软雅黑");
				break;
			case "SimSun":
				
				font.setFontname("宋体");
				break;
			case "SimHei":
				font.setFontname("黑体");
				break;
			case "KaiTi":
				
				font.setFontname("楷体");
				break;
			default:
				font.setFontname(content.getFamily());
				break;
			}
		}
		
		if (null == content.getSize()) {
			font.setSize((short) 220);
		} else {
			font.setSize((short) (Short.parseShort(content.getSize()) * 20));
		}

		if (null != content.getWeight() && content.getWeight()) {
			font.setBoldweight((short) 700);
		} else {
			font.setBoldweight((short) 400);
		}
		if (null == content.getColor()) {
			ExcelColor fontColor = new ExcelColor();
			font.setColor(fontColor);
		} else {
			ExcelColor fontColor = ExcelUtil.getColor(content.getColor());
			font.setColor(fontColor);
		}
		if (null == content.getUnderline()) {
			font.setUnderline((byte) 0);
		} else {
			font.setUnderline(content.getUnderline().byteValue());
		}
		if (null == content.getItalic()) {
			font.setItalic(false);
		} else {
			font.setItalic(content.getItalic());
		}
		style.setFont(font);
		if (null == content.getBackground()) {
			ExcelColor fgcolor = new ExcelColor(0, 0, 0);
			style.setFgcolor(fgcolor);
			ExcelColor bgcolor = new ExcelColor(0, 0, 0);
			style.setBgcolor(bgcolor);
		} else {
			ExcelColor fgcolor = ExcelUtil.getColor(content.getBackground());
			style.setFgcolor(fgcolor);
			ExcelColor bgcolor = new ExcelColor(0, 0, 0);
			style.setBgcolor(bgcolor);
			style.setPattern((short)1);
		}
		
        if(null == content.getExpress()){
        	style.setDataformat("General");
        }else{
        	style.setDataformat(content.getExpress());
        }
		
		if (null == content.getWordWrap()) {
			style.setWraptext(false);
		} else {
			style.setWraptext(content.getWordWrap());
		}
		if (null == content.getLocked()) {
			style.setLocked(true);
		} else {
			style.setLocked(content.getLocked());
		}

		cell.setCellstyle(style);

		return cell;
	}

	private ExcelRow getExcelRow(MRow mrow)  {
		ExcelRow row = new ExcelRow();

		Content content = mrow.getProps().getContent();
		Border border = mrow.getProps().getBorder();
		CustomProp cp = mrow.getProps().getCustomProp();

		row.setHeight(mrow.getHeight());
		try {
			row.setCode(mrow.getAlias());
		} catch (ExcelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if(null == mrow.getHidden()){
			row.setRowhidden(false);
		}else{
			row.setRowhidden(mrow.getHidden());
		}

		ExcelCellStyle style = new ExcelCellStyle();
		if(null == content.getAlignRow()){
			style.setAlign((short) 1);
		}else {
			switch (content.getAlignRow()) {
			case "left":
				style.setAlign((short) 1);
				break;
			case "center":
				style.setAlign((short) 2);
				break;
			case "right":
				style.setAlign((short) 3);
				break;
			default:
				style.setAlign((short) 2);
				break;
			}
		}
		
		if(null == content.getAlignCol()){
			style.setValign((short) 0);
		}else{
			switch (content.getAlignCol()) {
			case "top":
				style.setValign((short) 0);
				break;
			case "middle":
				style.setValign((short) 1);
				break;
			case "bottom":
				style.setValign((short) 2);
				break;
			default:
				style.setValign((short) 1);
				break;
			}
		}

		
		Excelborder topborder = new Excelborder();
		// ExcelColor topColor = ExcelUtil.getColor(border.getTop());
		if (null == border.getTop()) {
			topborder.setSort((short) 0);
		} else {
			topborder.setSort(border.getTop().shortValue());
		}
		style.setTopborder(topborder);
		Excelborder leftborder = new Excelborder();
		if (null == border.getLeft()) {
			leftborder.setSort((short) 0);
		} else {
			leftborder.setSort(border.getLeft().shortValue());
		}
		style.setLeftborder(leftborder);
		Excelborder rightborder = new Excelborder();
		if (null == border.getRight()) {
			rightborder.setSort((short) 0);
		} else {
			rightborder.setSort(border.getRight().shortValue());
		}
		style.setRightborder(rightborder);
		Excelborder bottomborder = new Excelborder();
		if (null == border.getBottom()) {
			bottomborder.setSort((short) 0);
		} else {
			bottomborder.setSort(border.getBottom().shortValue());
		}
		style.setBottomborder(bottomborder);

		ExcelFont font = new ExcelFont();
		if(null == content.getFamily()){
			font.setFontname("宋体");
		}else{
			
			switch (content.getFamily()) {
			case "microsoft Yahei":
				
				font.setFontname("微软雅黑");
				break;
			case "SimSun":
				
				font.setFontname("宋体");
				break;
			case "SimHei":
				font.setFontname("黑体");
				break;
			case "KaiTi":
				
				font.setFontname("楷体");
				break;
			default:
				font.setFontname(content.getFamily());
				break;
			}
		}
		if (null == content.getSize()) {
			font.setSize((short) 220);
		} else {
			font.setSize((short) (Short.parseShort(content.getSize()) * 20));
		}

		if (null != content.getWeight() && content.getWeight()) {
			font.setBoldweight((short) 700);
		} else {
			font.setBoldweight((short) 400);
		}
		if (null == content.getColor()) {
			ExcelColor fontColor = new ExcelColor();
			font.setColor(fontColor);
		} else {
			ExcelColor fontColor = ExcelUtil.getColor(content.getColor());
			font.setColor(fontColor);
		}
		if (null == content.getUnderline()) {
			font.setUnderline((byte) 0);
		} else {
			font.setUnderline(content.getUnderline().byteValue());
		}
		if (null == content.getItalic()) {
			font.setItalic(false);
		} else {
			font.setItalic(content.getItalic());
		}
		style.setFont(font);
		if (null == content.getBackground()) {
			ExcelColor fgcolor = new ExcelColor(0, 0, 0);
			style.setFgcolor(fgcolor);
		} else {
			ExcelColor fgcolor = ExcelUtil.getColor(content.getBackground());
			style.setFgcolor(fgcolor);
			style.setPattern((short)1);
		}

		if(null == content.getExpress()){
        	style.setDataformat("General");
        }else{
        	style.setDataformat(content.getExpress());
        }
		
		if (null == content.getWordWrap()) {
			style.setWraptext(false);
		} else {
			style.setWraptext(content.getWordWrap());
		}
		
		if (null == content.getLocked()) {
			style.setLocked(true);
		} else {
			style.setLocked(content.getLocked());
		}

		row.setCellstyle(style);

		return row;
	}

	private ExcelColumn getExcelCol(MCol mcol) {
		ExcelColumn col = new ExcelColumn();

		Content content = mcol.getProps().getContent();
		Border border = mcol.getProps().getBorder();
		CustomProp cp = mcol.getProps().getCustomProp();

		col.setWidth(mcol.getWidth());
		try {
			col.setCode(mcol.getAlias());
		} catch (ExcelException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		if(null == mcol.getHidden()){
			col.setColumnhidden(false);
		}else{
			col.setColumnhidden(mcol.getHidden());
		}
		
		ExcelCellStyle style = new ExcelCellStyle();
		if(null == content.getAlignRow()){
			style.setAlign((short) 1);
		}else {
			switch (content.getAlignRow()) {
			case "left":
				style.setAlign((short) 1);
				break;
			case "center":
				style.setAlign((short) 2);
				break;
			case "right":
				style.setAlign((short) 3);
				break;
			default:
				style.setAlign((short) 2);
				break;
			}
		}
		
		if(null == content.getAlignCol()){
			style.setValign((short) 0);
		}else{
			switch (content.getAlignCol()) {
			case "top":
				style.setValign((short) 0);
				break;
			case "middle":
				style.setValign((short) 1);
				break;
			case "bottom":
				style.setValign((short) 2);
				break;
			default:
				style.setValign((short) 1);
				break;
			}
		}
		Excelborder topborder = new Excelborder();
		// ExcelColor topColor = ExcelUtil.getColor(border.getTop());
		if (null == border.getTop()) {
			topborder.setSort((short) 0);
		} else {
			topborder.setSort(border.getTop().shortValue());
		}
		style.setTopborder(topborder);
		Excelborder leftborder = new Excelborder();
		if (null == border.getLeft()) {
			leftborder.setSort((short) 0);
		} else {
			leftborder.setSort(border.getLeft().shortValue());
		}
		style.setLeftborder(leftborder);
		Excelborder rightborder = new Excelborder();
		if (null == border.getRight()) {
			rightborder.setSort((short) 0);
		} else {
			rightborder.setSort(border.getRight().shortValue());
		}
		style.setRightborder(rightborder);
		Excelborder bottomborder = new Excelborder();
		if (null == border.getBottom()) {
			bottomborder.setSort((short) 0);
		} else {
			bottomborder.setSort(border.getBottom().shortValue());
		}
		style.setBottomborder(bottomborder);

		ExcelFont font = new ExcelFont();
		if(null == content.getFamily()){
			font.setFontname("宋体");
		}else{
			
			switch (content.getFamily()) {
			case "microsoft Yahei":
				
				font.setFontname("微软雅黑");
				break;
			case "SimSun":
				
				font.setFontname("宋体");
				break;
			case "SimHei":
				font.setFontname("黑体");
				break;
			case "KaiTi":
				
				font.setFontname("楷体");
				break;
			default:
				font.setFontname(content.getFamily());
				break;
			}
		}
		if (null == content.getSize()) {
			font.setSize((short) 220);
		} else {
			font.setSize((short) (Short.parseShort(content.getSize()) * 20));
		}

		if (null != content.getWeight() && content.getWeight()) {
			font.setBoldweight((short) 700);
		} else {
			font.setBoldweight((short) 400);
		}
		if (null == content.getColor()) {
			ExcelColor fontColor = new ExcelColor();
			font.setColor(fontColor);
		} else {
			ExcelColor fontColor = ExcelUtil.getColor(content.getColor());
			font.setColor(fontColor);
		}
		if (null == content.getUnderline()) {
			font.setUnderline((byte) 0);
		} else {
			font.setUnderline(content.getUnderline().byteValue());
		}
		if (null == content.getItalic()) {
			font.setItalic(false);
		} else {
			font.setItalic(content.getItalic());
		}
		style.setFont(font);
		if (null == content.getBackground()) {
			ExcelColor fgcolor = new ExcelColor(0, 0, 0);
			style.setFgcolor(fgcolor);
		} else {
			ExcelColor fgcolor = ExcelUtil.getColor(content.getBackground());
			style.setFgcolor(fgcolor);
			style.setPattern((short)1);
		}

		if(null == content.getExpress()){
        	style.setDataformat("General");
        }else{
        	style.setDataformat(content.getExpress());
        }
		
		if (null == content.getWordWrap()) {
			style.setWraptext(false);
		} else {
			style.setWraptext(content.getWordWrap());
		}
		if (null == content.getLocked()) {
			style.setLocked(true);
		} else {
			style.setLocked(content.getLocked());
		}

		col.setCellstyle(style);

		return col;
	}
	

	private void addRowOrCol(String excelId, String sheetId,
			List<RowCol> sortRList, List<RowCol> sortCList, int rowEnd,
			int colEnd) {
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		int maxRow = msheet.getMaxrow();
		int maxCol = msheet.getMaxcol();
		int rowHeight;
		if (sortRList.size() == 0) {
			rowHeight = 0;
		} else {
			RowCol rowTop = sortRList.get(sortRList.size() - 1);
			rowHeight = rowTop.getTop() + rowTop.getLength();
		}

		if (rowEnd > rowHeight) {
			int length = rowEnd - rowHeight;
			int rowNum = (length / 20) + 1;
			int top = rowHeight + 1;
			// 增加新的行
			for (int i = 0; i < rowNum; i++) {
				maxRow = maxRow + 1;
				String row = maxRow + "";
				MRow mrow = new MRow(row, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(row);
				rowCol.setLength(19);
				if (i == 0) {
					if (rowHeight == 0) {
						rowCol.setPreAlias(null);
					} else {
						rowCol.setPreAlias(
								sortRList.get(sortRList.size() - 1).getAlias());
					}

				} else {
					rowCol.setPreAlias((maxRow - 1) + "");
				}
				// 存入简化的行
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "rList");
				// 存入行样式
				baseDao.insert(excelId, mrow);

				if (i == 0) {
					rowCol.setTop(top);
				} else {
					top = top + 19 + 1;
					rowCol.setTop(top);
				}
				sortRList.add(rowCol);

			}
		}
		int colWeight;
		if (sortCList.size() == 0) {
			colWeight = 0;
		} else {
			RowCol colTop = sortCList.get(sortCList.size() - 1);
			colWeight = colTop.getTop() + colTop.getLength();
		}

		if (colEnd > colWeight) {
			int length = colEnd - colWeight;
			int colNum = (length / 72) + 1;
			int left = colWeight + 1;
			// 增加新的列
			for (int i = 0; i < colNum; i++) {
				maxCol = maxCol + 1;
				String col = maxCol + "";
				MCol mcol = new MCol(col, sheetId);
				RowCol rowCol = new RowCol();
				rowCol.setAlias(col);
				rowCol.setLength(71);
				if (i == 0) {
					if (colWeight == 0) {
						rowCol.setPreAlias(null);
					} else {
						rowCol.setPreAlias(
								sortCList.get(sortCList.size() - 1).getAlias());
					}
				} else {
					rowCol.setPreAlias((maxCol - 1) + "");
				}
				// 存入简化的列
				mrowColDao.insertRowCol(excelId, sheetId, rowCol, "cList");
				// 存入列样式
				baseDao.insert(excelId, mcol);
				if (i == 0) {
					rowCol.setTop(left);
				} else {
					left = left + 71 + 1;
					rowCol.setTop(left);
				}
				sortCList.add(rowCol);

			}

			msheet.setMaxcol(maxCol);
			msheet.setMaxrow(maxRow);
			baseDao.update(excelId, msheet);// 更新最大行，最大列
		}
	}

	private int getIndexByPixel(List<RowCol> sortRclist, int pixel) {
		if (sortRclist.size() == 0) {
			return 0;
		}

		RowCol rowColTop = sortRclist.get(sortRclist.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();

		int endPixel = 0;
		if (pixel == 0) {
			endPixel = 0;
		} else if (maxTop < pixel) {
			endPixel = maxTop;
		} else {
			endPixel = pixel;
		}
		int end = BinarySearch.binarySearch(sortRclist, endPixel);
		return end;
	}
	
	private int getIndexByPixelBegin(List<RowCol> sortRclist, int pixel) {
		if (sortRclist.size() == 0) {
			return 0;
		}

		RowCol rowColTop = sortRclist.get(sortRclist.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();

		int endPixel = 0;
		if (pixel == 0) {
			endPixel = 0;
		} else if (maxTop < pixel) {
			return -1;
		} else {
			endPixel = pixel;
		}
		int end = BinarySearch.binarySearch(sortRclist, endPixel);
		return end;
	}

	private int getColLeft(List<Glx> glxList, int i, Integer left) {
		if (i == 0) {
			return left;
		}
		Glx glx = glxList.get(i - 1);
		int tempWidth = glx.getWidth();
		if (null != glx.getHidden() && glx.getHidden()) {
			tempWidth = 0;
		}
		return glx.getLeft() + tempWidth + 1;
	}

	private int getRowTop(List<Gly> glyList, int i, int top) {
		if (i == 0) {
			return top;
		}
		Gly gly = glyList.get(i - 1);
		int tempHeight = gly.getHeight();
		if (null != gly.getHidden() && gly.getHidden()) {
			tempHeight = 0;
		}

		return gly.getTop() + tempHeight + 1;
	}

	@Override
	public List<String> getExcels() {

		return msheetDao.getTableList();
	}

}
