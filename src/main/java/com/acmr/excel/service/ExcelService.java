package com.acmr.excel.service;



import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.ExcelDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OnlineExcel;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.util.CellFormateUtil;
import com.acmr.excel.util.ExcelUtil;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.Excelborder;



/**
 * excel的还原操作service
 * 
 * @author jinhr
 *
 */
@Service
public class ExcelService {
	private ExcelDao excelDao;

	public void setExcelDao(ExcelDao excelDao) {
		this.excelDao = excelDao;
	}

	/**
	 * 通过像素还原excel
	 * 
	 * @param rowBegin
	 *            开始行像素
	 * @param rowEnd
	 *            开始列像素
	 * @return CompleteExcel对象
	 */

	public SheetElement openExcel(SheetElement spreadSheet,ExcelSheet excelSheet, int rowBeginIndex, int rowEndIndex,int colBeginIndex,int colEndIndex,
			List<RowCol> rList,List<RowCol> cList) {
		List<Gly> glyList = spreadSheet.getGridLineRow();
		List<Glx> glxList = spreadSheet.getGridLineCol();
		for(int i = rowBeginIndex;i<= rowEndIndex;i++) {
			Gly gly = new Gly();
			gly.setAlias(rList.get(i).getAlias());
			gly.setHeight(rList.get(i).getLength());
			gly.setTop(rList.get(i).getTop());
			gly.setSort(i);
			glyList.add(gly);
		}
     
		/*Gly gly = glyList.get(glyList.size()-1);
		if (rowEnd > gly.getTop() + gly.getHeight()) {
			addRow(glyList, rowEnd, excelSheet);
		}*/
		//bookToOlExcelGlxList(excelSheet, glxList);
		for (int i = colBeginIndex; i <= colEndIndex; i++) {
			Glx glx = new Glx();
			glx.setAlias(cList.get(i).getAlias());
			glx.setWidth(cList.get(i).getLength());
			glx.setLeft(cList.get(i).getTop());
			glx.setSort(i);
			glxList.add(glx);
		}
		List<ExcelRow> rowList = excelSheet.getRows();

		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.setCells(newCellList);
		bookToOlExcelCellList2(rowList, glyList,glxList, excelSheet, newCellList,rowBeginIndex,colBeginIndex);
		//spreadSheet.getSheet().setGlY(glyList.subList(rowBeginIndex, rowEndIndex + 1));
	
		
		return spreadSheet;
	}
	
	
	private void addRow(List<Gly> glyList, int rowEnd, ExcelSheet excelSheet) {
		int index = glyList.size() - 1;
		Gly gly = glyList.get(index);
		int top = gly.getTop();
		int length = rowEnd - top - gly.getHeight();
		int rowNum = (length / 20) + 1;
		for (int i = 0; i < rowNum; i++) {
			ExcelRow row = excelSheet.addRow();
			int tmp = glyList.size();
			Gly newGly = new Gly();
			newGly.setSort(tmp);
			newGly.setHeight(19);
			newGly.setAlias(row.getCode());
			newGly.setTop(getRowTop(glyList, tmp++));
			glyList.add(newGly);
		}
		
		
	}
	
	private void addEmptyRow(List<Gly> glyList, int height, ExcelSheet excelSheet) {
		int rowNum = (height / 20) + 1;
		for (int i = 0; i < rowNum; i++) {
			ExcelRow row = excelSheet.addRow();
			int tmp = glyList.size();
			Gly newGly = new Gly();
			newGly.setSort(tmp);
			newGly.setHeight(19);
			newGly.setAlias(row.getCode());
			newGly.setTop(getRowTop(glyList, tmp++));
			glyList.add(newGly);
		}
		
		
	}

	/**
	 * 通过别名加载excel
	 * 
	 * @return CompleteExcel对象
	 */

	public SpreadSheet openExcelByAlais(SpreadSheet spreadSheet,
			ExcelSheet excelSheet, String rowBeginAlais, String rowEndAlais,
			ReturnParam returnParam) {
//		List<Gly> glyList = spreadSheet.getSheet().getGlY();
//		List<Glx> glxList = spreadSheet.getSheet().getGlX();
//		bookToOlExcelGlyList(excelSheet, glyList);
//		bookToOlExcelGlxList(excelSheet, glxList);
//		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet
//				.getRows();
//		int rowBeginIndex = rowList.getMaps().get(rowBeginAlais);
//		int rowEndIndex = rowList.getMaps().get(rowEndAlais);
//		List<OneCell> newCellList = new ArrayList<OneCell>();
//		spreadSheet.getSheet().setCells(newCellList);
//		bookToOlExcelCellList(rowBeginIndex, rowEndIndex, rowList, glyList,glxList, excelSheet, newCellList);
//		spreadSheet.getSheet().setGlY(glyList.subList(rowBeginIndex, rowEndIndex + 1));
//		returnParam.setDataRowStartIndex(rowBeginIndex);
		return spreadSheet;
	}

	/**
	 * 通过Workbook转换为CompleteExcel
	 * 
	 * @param excel
	 *            CompleteExcel对象
	 * @return CompleteExcel对象
	 */

//	public CompleteExcel getExcel(Workbook workbook) {
//		Convert convert = new OnlineExcelConvert();
//		return convert.doConvertExcel(workbook);
//	}

	/**
	 * 保存excel
	 * 
	 * @param excel
	 *            OnlineExcel对象
	 */

	public void saveOrUpdateExcel(OnlineExcel excel) throws Exception {
		String excelId = excel.getExcelId();
		if(excelDao.getByExcelId(excelId) == 0){
			excelDao.saveExcel(excel);
		}
	}

	/**
	 * 获得所有的OnlineExcel对象
	 * 
	 * @return OnlineExcel对象集合
	 */

	public List<OnlineExcel> getAllExcel() {
		return excelDao.getAllExcel();
	}

	

	private int getColLeft(List<Glx> glxList, int i) {
		if (i == 0) {
			return 0;
		}
		Glx glx = glxList.get(i - 1);
		int tempWidth = glx.getWidth();
		if(null!=glx.getHidden()&&glx.getHidden()){
			tempWidth = -1;
		}
		return glx.getLeft() + tempWidth + 1;
	}

	private int getRowTop(List<Gly> glyList, int i) {
		if (i == 0) {
			return 0;
		}
		Gly gly = glyList.get(i - 1);
		int tempHeight = gly.getHeight();

		return gly.getTop() + tempHeight + 1;
	}
	

	private void bookToOlExcelCellList(int rowBeginIndex, int rowEndIndex,
			List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,
			ExcelSheet excelSheet, List<OneCell> newCellList) {
		getCellList(rowBeginIndex, rowEndIndex, rowList, glyList, glxList, excelSheet, newCellList);
	}
	
	private void bookToOlExcelGlyList(ExcelSheet excelSheet, List<Gly> glyList) {
		List<ExcelRow> rowList = excelSheet.getRows();
		for (int i = 0; i < rowList.size(); i++) {
			ExcelRow excelRow = rowList.get(i);
			Gly gly = new Gly();
			gly.setAlias(rowList.get(i).getCode());
			gly.setHeight(excelRow.getHeight());
            gly.setHidden(excelRow.isRowhidden());
			gly.setTop(getRowTop(glyList, i));
			gly.setSort(i);
			glyList.add(gly);
		}

	}

	private void bookToOlExcelGlxList(ExcelSheet excelSheet, List<Glx> glxList) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for (int i = 0; i < colList.size(); i++) {
			ExcelColumn excelColumn = colList.get(i);
			ExcelCellStyle style = excelColumn.getCellstyle();
			ExcelFont font = style.getFont();
			
			Glx glx = new Glx();
			glx.setAlias(excelColumn.getCode());
			
			boolean hidden = excelColumn.isColumnhidden();
			if(hidden){
				glx.setHidden(true);
				glx.setWidth(0);
			}else{
				glx.setWidth(excelColumn.getWidth());
				glx.setHidden(false);
			}
			
			glx.setLeft(getColLeft(glxList, i));
			glx.setSort(i);
			
			OperProp operProp = glx.getProps();
			
			Content content = operProp.getContent();
			switch (style.getAlign()) {
			case 0:
				content.setAlignRow(null);
				break;
			case 1:
				content.setAlignRow("center");
				break;
			case 2:
				content.setAlignRow("right");
				break;

			default:
				
				break;
			}
			switch (style.getValign()) {
			case 1:
				content.setAlignCol(null);
				break;
			case 2:
				content.setAlignCol("middle");
				break;
			case 3:
				content.setAlignCol("bottom");
				break;

			default:
				
				break;
			}
			
			if (700 == font.getBoldweight()) {
				content.setBold(true);
			} else {
				content.setBold(null);
			}
			if(null==font.getColor()){
				content.setColor(null);
			}else{
				content.setColor(CellFormateUtil.getBackground(font.getColor()));
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
			if(font.isItalic()){
				content.setItalic(true);
			}else{
				content.setItalic(null);
			}
			if(font.getSize()==200){
				content.setSize(null);
			}else{
				content.setSize(font.getSize() / 20 + "");
			}
			content.setTexts(null);
			if(0==(int)font.getUnderline()){
				content.setUnderline(null);
			}else{
				content.setUnderline((int)font.getUnderline());
			}
			if(style.isWraptext()){
				content.setWordWrap(style.isWraptext());
			}else{
				content.setWordWrap(null);
			}
			content.setType(null);
			content.setExpress(null);
			if(null==style.getFgcolor()){
				content.setBackground(null);
			}else{
				content.setBackground(CellFormateUtil.getBackground(style.getFgcolor()));
			}
			if(style.isLocked()){
				content.setLocked(null);
			}else{
				content.setLocked(style.isLocked());
				
			}
			

			CustomProp customProp = operProp.getCustomProp();
			customProp.setComment(null);
			
			Border border = operProp.getBorder();
			Excelborder excelTopborder = style.getTopborder();
			if(excelTopborder != null){
				short topBorder = excelTopborder.getSort();
				if (topBorder == 2) {
					border.setTop(2);
				}else if(topBorder == 1){
					border.setTop(1);
				}else{
					border.setTop(null);
				}
			}
			Excelborder excelBottomborder  = style.getBottomborder();
			if(excelBottomborder != null){
				short bottomBorder = excelBottomborder.getSort();
				if (bottomBorder == 2) {
					border.setBottom(2);
				}else if(bottomBorder == 1){
					border.setBottom(1);
				}else{
					border.setBottom(null);
				}
			}
			Excelborder excelLeftborder = style.getLeftborder();
			if(excelLeftborder != null){
				short leftBorder = excelLeftborder.getSort();
				if (leftBorder == 2) {
					border.setLeft(2);
				}else if(leftBorder == 1){
					border.setLeft(1);
				}else{
					border.setLeft(null);
				}
			}
			Excelborder excelRightborder = style.getRightborder();
			if(excelRightborder != null){
				short rightBorder = excelRightborder.getSort();
				if (rightBorder == 2) {
					border.setRight(2);
				}else if(rightBorder == 1){
					border.setRight(1);
				}else{
					border.setRight(null);
				}
			}
			
			operProp.setContent(content);
			operProp.setCustomProp(customProp);
			operProp.setBorder(border);
			glx.setProps(operProp);
			glxList.add(glx);
		}
	}
	
	private void bookToOlExcelCellList2(List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,ExcelSheet excelSheet, List<OneCell> newCellList
			,int rowBeginIndex,int colBeginIndex) {
		for (int i = 0; i < rowList.size(); i++) {
			rowBeginIndex = i+rowBeginIndex;
			ExcelRow excelRow = rowList.get(i);
			if (excelRow != null) {
				List<ExcelCell> cellList = excelRow.getCells();
				
				for (int j = 0; j < cellList.size(); j++) {
					colBeginIndex = j+colBeginIndex;
					ExcelCell excelCell = cellList.get(j);
					if (excelCell != null) {
						OneCell cell = new OneCell();
					
						setCellProperty(excelCell, cell, i, j, glyList,glxList, excelSheet);

						newCellList.add(cell);
					}
				}
				ExcelCellStyle style = excelRow.getCellstyle();
				ExcelFont   font = style.getFont();
				OperProp   prop = glyList.get(i).getProps();
				
				Content content = prop.getContent();
				
				switch (style.getAlign()) {
				case 0:
					content.setAlignRow(null);
					break;
				case 1:
					content.setAlignRow("center");
					break;
				case 2:
					content.setAlignRow("right");
					break;

				default:
					
					break;
				}
				switch (style.getValign()) {
				case 1:
					content.setAlignCol(null);
					break;
				case 2:
					content.setAlignCol("middle");
					break;
				case 3:
					content.setAlignCol("bottom");
					break;

				default:
					
					break;
				}
				
				if (700 == font.getBoldweight()) {
					content.setBold(true);
				} else {
					content.setBold(null);
				}
				if(null==font.getColor()){
					content.setColor(null);
				}else{
					content.setColor(CellFormateUtil.getBackground(font.getColor()));
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
				if(font.isItalic()){
					content.setItalic(true);
				}else{
					content.setItalic(null);
				}
				if(font.getSize()==220){
					content.setSize(null);
				}else{
					content.setSize(font.getSize() / 20 + "");
				}
				content.setTexts(null);
				if(0==(int)font.getUnderline()){
					content.setUnderline(null);
				}else{
					content.setUnderline((int)font.getUnderline());
				}
				if(style.isWraptext()){
					content.setWordWrap(style.isWraptext());
				}else{
					content.setWordWrap(null);
				}
				content.setType(null);
				content.setExpress(null);
				if(null==style.getFgcolor()){
					content.setBackground(null);
				}else{
					content.setBackground(CellFormateUtil.getBackground(style.getFgcolor()));
				}
				if(style.isLocked()){
					content.setLocked(null);
				}else{
					
					content.setLocked(style.isLocked());
				}
				

				CustomProp customProp = prop.getCustomProp();
				customProp.setComment(null);
				
				Border border = prop.getBorder();
				Excelborder excelTopborder = style.getTopborder();
				if(excelTopborder != null){
					short topBorder = excelTopborder.getSort();
					if (topBorder == 2) {
						border.setTop(2);
					}else if(topBorder == 1){
						border.setTop(1);
					}else{
						border.setTop(null);
					}
				}
				Excelborder excelBottomborder  = style.getBottomborder();
				if(excelBottomborder != null){
					short bottomBorder = excelBottomborder.getSort();
					if (bottomBorder == 2) {
						border.setBottom(2);
					}else if(bottomBorder == 1){
						border.setBottom(1);
					}else{
						border.setBottom(null);
					}
				}
				Excelborder excelLeftborder = style.getLeftborder();
				if(excelLeftborder != null){
					short leftBorder = excelLeftborder.getSort();
					if (leftBorder == 2) {
						border.setLeft(2);
					}else if(leftBorder == 1){
						border.setLeft(1);
					}else{
						border.setLeft(null);
					}
				}
				Excelborder excelRightborder = style.getRightborder();
				if(excelRightborder != null){
					short rightBorder = excelRightborder.getSort();
					if (rightBorder == 2) {
						border.setRight(2);
					}else if(rightBorder == 1){
						border.setRight(1);
					}else{
						border.setRight(null);
					}
				}
					
			}
		}
	}

	/**
	 * 获得普通单元格
	 * 
	 * @param rowBeginIndex
	 * @param rowEndIndex
	 * @param rowList
	 * @param glyList
	 * @param glxList
	 * @param excelSheet
	 * @param newCellList
	 */

	private void getCellList(int rowBeginIndex, int rowEndIndex,List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,
			ExcelSheet excelSheet, List<OneCell> newCellList) {
		for (int i = rowBeginIndex; i <= rowEndIndex; i++) {
			ExcelRow excelRow = rowList.get(i);
			if (excelRow != null) {
				List<ExcelCell> cellList = excelRow.getCells();
				for (int j = 0; j <= glxList.size()-1; j++) {
					ExcelCell excelCell = cellList.get(j);
					if (excelCell != null) {
						OneCell cell = new OneCell();
					/*	String highLight = excelCell.getExps().get(Constant.HIGHLIGHT);
						if ("true".equals(highLight)) {
							cell.setHighlight(true);
						} else {
							cell.setHighlight(false);
						}*/
						setCellProperty(excelCell, cell, i, j, glyList,glxList, excelSheet);
						if (excelSheet.getCols().get(j).isColumnhidden() || excelSheet.getRows().get(i).isRowhidden()) {
							int colspan = excelCell.getColspan();
							int rowspan = excelCell.getRowspan();
							if (colspan > 1 || rowspan > 1) {
								int[] mfCell = excelSheet.getMergFirstCell(i, j);

								int tempX = mfCell[1];
								boolean flag = true;

								for (int k = tempX; k < tempX + colspan; k++) {
									if (!excelSheet.getCols().get(k).isColumnhidden()) {
										flag = false;
									}
								}
								
							}
						}
						newCellList.add(cell);
					}
				}
				
				ExcelCellStyle style = excelRow.getCellstyle();
				ExcelFont   font = style.getFont();
				OperProp   prop = glyList.get(i).getProps();
				
				Content content = prop.getContent();
				
				switch (style.getAlign()) {
				case 0:
					content.setAlignRow(null);
					break;
				case 1:
					content.setAlignRow("center");
					break;
				case 2:
					content.setAlignRow("right");
					break;

				default:
					
					break;
				}
				switch (style.getValign()) {
				case 1:
					content.setAlignCol(null);
					break;
				case 2:
					content.setAlignCol("middle");
					break;
				case 3:
					content.setAlignCol("bottom");
					break;

				default:
					
					break;
				}
				
				if (700 == font.getBoldweight()) {
					content.setBold(true);
				} else {
					content.setBold(null);
				}
				if(null==font.getColor()){
					content.setColor(null);
				}else{
					content.setColor(CellFormateUtil.getBackground(font.getColor()));
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
				if(font.isItalic()){
					content.setItalic(true);
				}else{
					content.setItalic(null);
				}
				if(font.getSize()==220){
					content.setSize(null);
				}else{
					content.setSize(font.getSize() / 20 + "");
				}
				content.setTexts(null);
				if(0==(int)font.getUnderline()){
					content.setUnderline(null);
				}else{
					content.setUnderline((int)font.getUnderline());
				}
				if(style.isWraptext()){
					content.setWordWrap(style.isWraptext());
				}else{
					content.setWordWrap(null);
				}
				content.setType(null);
				content.setExpress(null);
				if(null==style.getFgcolor()){
					content.setBackground(null);
				}else{
					content.setBackground(CellFormateUtil.getBackground(style.getFgcolor()));
				}
				if(style.isLocked()){
					content.setLocked(null);
				}else{
					
					content.setLocked(style.isLocked());
				}
				

				CustomProp customProp = prop.getCustomProp();
				customProp.setComment(null);
				
				Border border = prop.getBorder();
				Excelborder excelTopborder = style.getTopborder();
				if(excelTopborder != null){
					short topBorder = excelTopborder.getSort();
					if (topBorder == 2) {
						border.setTop(2);
					}else if(topBorder == 1){
						border.setTop(1);
					}else{
						border.setTop(null);
					}
				}
				Excelborder excelBottomborder  = style.getBottomborder();
				if(excelBottomborder != null){
					short bottomBorder = excelBottomborder.getSort();
					if (bottomBorder == 2) {
						border.setBottom(2);
					}else if(bottomBorder == 1){
						border.setBottom(1);
					}else{
						border.setBottom(null);
					}
				}
				Excelborder excelLeftborder = style.getLeftborder();
				if(excelLeftborder != null){
					short leftBorder = excelLeftborder.getSort();
					if (leftBorder == 2) {
						border.setLeft(2);
					}else if(leftBorder == 1){
						border.setLeft(1);
					}else{
						border.setLeft(null);
					}
				}
				Excelborder excelRightborder = style.getRightborder();
				if(excelRightborder != null){
					short rightBorder = excelRightborder.getSort();
					if (rightBorder == 2) {
						border.setRight(2);
					}else if(rightBorder == 1){
						border.setRight(1);
					}else{
						border.setRight(null);
					}
				}
					
			}
		}
	}


	/**
	 * 设置单元格常规属性
	 * 
	 * @param excelCell
	 * @param cell
	 * @param i
	 * @param j
	 * @param glyList
	 * @param glxList
	 * @param excelSheet
	 */

	private void setCellProperty(ExcelCell excelCell, OneCell cell, int i,int j, List<Gly> glyList, List<Glx> glxList, ExcelSheet excelSheet) {
		ExcelCellStyle excelCellStyle = excelCell.getCellstyle();
		Border border = cell.getBorder();
		Excelborder excelTopborder = excelCellStyle.getTopborder();
		if(excelTopborder != null){
			short topBorder = excelTopborder.getSort();
			if (topBorder == 2) {
				border.setTop(2);
			}else if(topBorder == 1){
				border.setTop(1);
			}else{
				border.setTop(null);
			}
		}
		Excelborder excelBottomborder  = excelCellStyle.getBottomborder();
		if(excelBottomborder != null){
			short bottomBorder = excelBottomborder.getSort();
			if (bottomBorder == 2) {
				border.setBottom(2);
			}else if(bottomBorder == 1){
				border.setBottom(1);
			}else{
				border.setBottom(null);
			}
		}
		Excelborder excelLeftborder = excelCellStyle.getLeftborder();
		if(excelLeftborder != null){
			short leftBorder = excelLeftborder.getSort();
			if (leftBorder == 2) {
				border.setLeft(2);
			}else if(leftBorder == 1){
				border.setLeft(1);
			}else{
				border.setLeft(null);
			}
		}
		Excelborder excelRightborder = excelCellStyle.getRightborder();
		if(excelRightborder != null){
			short rightBorder = excelRightborder.getSort();
			if (rightBorder == 2) {
				border.setRight(2);
			}else if(rightBorder == 1){
				border.setRight(1);
			}else{
				border.setRight(null);
			}
		}
		
		Content content = cell.getContent();
		switch (excelCellStyle.getAlign()) {
		case 0:
			content.setAlignRow("left");
			break;
		case 1:
			content.setAlignRow("center");
			break;
		case 2:
			content.setAlignRow("right");
			break;

		default:
			
			break;
		}
		switch (excelCellStyle.getValign()) {
		case 1:
			content.setAlignCol("top");
			break;
		case 2:
			content.setAlignCol("middle");
			break;
		case 3:
			content.setAlignCol("bottom");
			break;

		default:
			
			break;
		}
		ExcelFont excelFont = excelCellStyle.getFont();
		switch (excelFont.getFontname()) {
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
			content.setFamily(excelFont.getFontname());
			break;
		}
		if (700 == excelFont.getBoldweight()) {
			content.setBold(true);
		} else {
			content.setBold(false);
		}
		content.setItalic(excelFont.isItalic());
		content.setSize(excelFont.getSize() / 20 + "");
		content.setTexts(excelCell.getText());
		
		ExcelColor fontColor = excelFont.getColor();
		if (fontColor != null) {
			int[] rgb = fontColor.getRGBInt();
			content.setColor(ExcelUtil.getRGB(rgb));
			
		}
		
		CellFormateUtil.setShowText(excelCell, content);
		content.setUnderline((int)excelFont.getUnderline());
		content.setWordWrap(excelCellStyle.isWraptext());
		content.setType(excelCell.getText());
		content.setExpress(null);
		ExcelColor color = excelCellStyle.getFgcolor();
		if (color != null) {
			int[] rgb = color.getRGBInt();
			content.setBackground(ExcelUtil.getRGB(rgb));
		}else{
			content.setBackground(null);
		}
		
		content.setLocked(excelCellStyle.isLocked());
		
		CustomProp customProp = cell.getCustomProp();
		customProp.setComment(excelCell.getMemo());
		String highLight = excelCell.getExps().get(Constant.HIGHLIGHT);
		if ("true".equals(highLight)) {
			customProp.setHighlight(true);
		} else {
			customProp.setHighlight(false);
		}
	
		Occupy occupy = cell.getOccupy();
		int rowspan = excelCell.getRowspan();
		int colspan = excelCell.getColspan();
		int[] firstMergeCell = excelSheet.getMergFirstCell(i, j);
		if (firstMergeCell == null) {
			occupy.getCol().add(glxList.get(j).getAlias());
			occupy.getRow().add(glyList.get(i).getAlias());
		
		} else {
			int firstRow = firstMergeCell[0];
			int firstCol = firstMergeCell[1];
			
			for (int m = firstRow; m < firstRow + rowspan; m++) {
				occupy.getRow().add(glyList.get(m).getAlias());
			}
			for (int m = firstCol; m < firstCol + colspan; m++) {

				occupy.getCol().add(glxList.get(m).getAlias());
			}
		}
		
	
	}
	
	
	
	/**
	 * 非冻结定位还原excel
	 * 
	 * @param height
	 *            高度
	 * @param returnParam
	 *            返回参数
	 * @return
	 */

	public SheetElement positionExcel(ExcelSheet excelSheet, SheetElement spreadSheet, int height,int end) {
		spreadSheet.setName(excelSheet.getName());
		List<Gly> glyList = spreadSheet.getGridLineRow();
		List<Glx> glxList = spreadSheet.getGridLineCol();
		bookToOlExcelGlyList(excelSheet, glyList);
		if(glyList.size() ==0){
			addEmptyRow(glyList, height, excelSheet);
		}
		Gly gly = glyList.get(glyList.size()-1);
		if (height > gly.getTop() + gly.getHeight()) {
			addRow(glyList, height, excelSheet);
		}
		bookToOlExcelGlxList(excelSheet, glxList);
		List<ExcelRow> rowList = excelSheet.getRows();
		// List<ExcelColumn> colList = excelSheet.getCols();
		int minTop = glyList.get(0).getTop();
		Gly glyTop = glyList.get(glyList.size() - 1);
		int maxTop = glyTop.getTop() + glyTop.getHeight();
		
		spreadSheet.setMaxRowPixel(maxTop);
		spreadSheet.setMaxRowAlias(glyTop.getAlias());
		
		int startAlaisPixel = 0;
		int Offset = startAlaisPixel - minTop;
		int startPixel = Offset < 200 ? Offset : startAlaisPixel - 200;
		int endPixel = 0;
		if (maxTop < height) {
			endPixel = maxTop;
		} else {
			endPixel = startPixel + height + 200;
		}
		
		Glx glxLeft = glxList.get(glxList.size() - 1);
		int maxLeft = glxLeft.getLeft() + glxLeft.getWidth();
		
		spreadSheet.setMaxColPixel(maxLeft);
		spreadSheet.setMaxColAlias(glxLeft.getAlias());
		
		int rowBeginIndex = 0;
		
		int rowEndIndex = end;
		
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.setCells(newCellList);
		bookToOlExcelCellList(rowBeginIndex, rowEndIndex,rowList, glyList,glxList, excelSheet, newCellList);
		spreadSheet.setGridLineRow(glyList.subList(rowBeginIndex, rowEndIndex + 1));
		
		/*ExcelSheetFreeze excelfrozen = excelSheet.getFreeze();
		if (excelfrozen != null) {
			Frozen frozen = spreadSheet.getFrozen();
			int rf = excelfrozen.getRow();
			if("fr".equals(excelSheet.getExps().get("fr"))){
				frozen.setRow(-1);
			}else{
				frozen.setRow(rf);
			}
			int cf = excelfrozen.getCol();
			if("fc".equals(excelSheet.getExps().get("fc"))){
				frozen.setCol(-1);
			}else{
				frozen.setCol(cf);
			}

		}*/
		return spreadSheet;
	}


	/**
	 * 通过id获得OnlineExcel对象
	 * 
	 * @param id
	 * @return OnlineExcel
	 */

	public String getExcel(String excelId) {
		try {
			OnlineExcel oe = excelDao.getJsonObjectByExcelId(excelId);
			return oe.getExcelObject();
		} catch (SQLException e) {
			e.printStackTrace();
		}
		return null;
	}


	/**
	 * 改变宽度或高度
	 * 
	 * @param excelBook
	 */

	public void changeHeightOrWidth(ExcelBook excelBook) {
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		List<ExcelColumn> colList = excelSheet.getCols();
		List<ExcelRow> rowList = excelSheet.getRows();
		for (ExcelColumn excelColumn : colList) {
			excelColumn.setWidth(excelColumn.getWidth() * 40);
		}
		for (ExcelRow excelRow : rowList) {
			excelRow.setHeight(excelRow.getHeight() * 18);
		}
	}
}
