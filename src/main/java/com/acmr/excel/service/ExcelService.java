package com.acmr.excel.service;



import java.sql.SQLException;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Service;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelFormat;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.ExcelSheetFreeze;
import acmr.excel.pojo.Excelborder;

import com.acmr.excel.dao.ExcelDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OnlineExcel;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Format;
import com.acmr.excel.model.complete.Frozen;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.CellFormateUtil;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.StringUtil;



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

	public SpreadSheet openExcel(SpreadSheet spreadSheet,ExcelSheet excelSheet, int rowBegin, int rowEnd,int colBegin,int colEnd,ReturnParam returnParam,
			List<RowCol> rList,List<RowCol> cList) {
		List<Gly> glyList = spreadSheet.getSheet().getGlY();
		List<Glx> glxList = spreadSheet.getSheet().getGlX();
		//bookToOlExcelGlyList(excelSheet, glyList);
		for(int i = rowBegin;i<= rowEnd;i++) {
			Gly gly = new Gly();
			gly.setAliasY(rList.get(i).getAlias());
//			gly.setOriginHeight(rowList.get(i).getHeight());
//			boolean hidden = excelRow.isRowhidden();
//			if(hidden){
//				gly.setHidden(true);
//				gly.setHeight(0);
//			}else{
				gly.setHeight(rList.get(i).getLength());
//			}
			gly.setTop(rList.get(i).getTop());
			gly.setIndex(i);
			glyList.add(gly);
		}
		Gly gly = glyList.get(glyList.size()-1);
		if (rowEnd > gly.getTop() + gly.getHeight()) {
			addRow(glyList, rowEnd, excelSheet);
		}
		bookToOlExcelGlxList(excelSheet, glxList);
		List<ExcelRow> rowList = excelSheet.getRows();
//		int rowBeginIndex = rowBegin;
//		int rowEndIndex = rowEnd;
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.getSheet().setCells(newCellList);
		bookToOlExcelCellList(0, rowList.size()-1,rowList, glyList,glxList, excelSheet, newCellList);
		//spreadSheet.getSheet().setGlY(glyList.subList(rowBeginIndex, rowEndIndex + 1));
		returnParam.setDataRowStartIndex(rowBegin);
		Gly maxGly = glyList.get(glyList.size()-1);
		returnParam.setMaxRowPixel(maxGly.getTop()+maxGly.getHeight());
		return spreadSheet;
	}
	
	
	private void addRow(List<Gly> glyList, int rowEnd, ExcelSheet excelSheet) {
		int index = glyList.size() - 1;
		Gly gly = glyList.get(index);
		int top = gly.getTop();
		int length = rowEnd - top - 19;
		int rowNum = (length / 20) + 1;
		for (int i = 0; i < rowNum; i++) {
			ExcelRow row = excelSheet.addRow();
			int tmp = glyList.size();
			Gly newGly = new Gly();
			newGly.setIndex(tmp);
			newGly.setHeight(19);
			newGly.setAliasY(row.getCode());
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
			newGly.setIndex(tmp);
			newGly.setHeight(19);
			newGly.setAliasY(row.getCode());
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

	private void bookToOlExcelGlyList(ExcelSheet excelSheet, List<Gly> glyList) {
		List<ExcelRow> rowList = excelSheet.getRows();
		for (int i = 0; i < rowList.size(); i++) {
			ExcelRow excelRow = rowList.get(i);
			Gly gly = new Gly();
			gly.setAliasY(rowList.get(i).getCode());
//			gly.setOriginHeight(rowList.get(i).getHeight());
//			boolean hidden = excelRow.isRowhidden();
//			if(hidden){
//				gly.setHidden(true);
//				gly.setHeight(0);
//			}else{
				gly.setHeight(excelRow.getHeight());
//			}
			gly.setTop(getRowTop(glyList, i));
			gly.setIndex(i);
			glyList.add(gly);
		}

	}

	private void bookToOlExcelGlxList(ExcelSheet excelSheet, List<Glx> glxList) {
		List<ExcelColumn> colList = excelSheet.getCols();
		for (int i = 0; i < colList.size(); i++) {
			ExcelColumn excelColumn = colList.get(i);
			Glx glx = new Glx();
			glx.setAliasX(excelColumn.getCode());
			glx.setOriginWidth(excelColumn.getWidth());
			boolean hidden = excelColumn.isColumnhidden();
			if(hidden){
				glx.setHidden(true);
				glx.setWidth(0);
			}else{
				glx.setWidth(excelColumn.getWidth());
			}
			
			glx.setLeft(getColLeft(glxList, i));
			glx.setIndex(i);
			Map<String, String> colMap = excelColumn.getExps();
			//Content content = glxList.get(i).getOperProp().getContent();
			OperProp operProp = glx.getOperProp();
			Content content = operProp.getContent();
			String alignCol = colMap.get("align_vertical");
			content.setAlignCol(alignCol);
			String alignRow = colMap.get("align_level");
			content.setAlignRow(alignRow);
			String bd = colMap.get("font_weight");
			if (bd != null) {
				content.setBd(Boolean.valueOf(bd));
			} else {
				content.setBd(null);
			}
			String color = colMap.get("font_color");
			content.setColor(color);
			content.setFamily(colMap.get("font_family"));
			String italic = colMap.get("font_italic");
			if (italic != null) {
				content.setItalic(Boolean.valueOf(italic));
			} else {
				content.setItalic(null);
			}
			content.setSize(colMap.get("font_size"));
			content.setRgbColor(null);
			content.setTexts(null);
			content.setAlignLine(null);
//			CustomProp customProp = glxList.get(i).getOperProp()
//					.getCustomProp();
			CustomProp customProp = operProp.getCustomProp();
			Format formate = operProp.getFormate();
			customProp.setBackground(colMap.get("fill_bgcolor"));
			formate.setType(colMap.get("format")); 
			formate.setCurrencySign(colMap.get("currency"));
			formate.setDateFormat(colMap.get("dateFormat"));
			customProp.setIsValid(null);
			String decimalPoint = colMap.get("decimalPoint");
			if (decimalPoint != null) {
				formate.setDecimal(Integer.valueOf(decimalPoint));
			} else {
				formate.setDecimal(null);
			}
			String thousandPoint = colMap.get("thousandPoint");
			if (thousandPoint != null) {
				formate.setThousands(Boolean.valueOf(thousandPoint));
			} else {
				formate.setThousands(null);
			}

			customProp.setComment(colMap.get("comment"));
			//Border border = glxList.get(i).getOperProp().getBorder();
			Border border = operProp.getBorder();
			String bottom = colMap.get("bottom");
			String top = colMap.get("top");
			String left = colMap.get("left");
			String right = colMap.get("right");
			String all = colMap.get("all");
			String outer = colMap.get("outer");
			String none = colMap.get("none");
			if (bottom != null) {
				border.setBottom(Boolean.valueOf(bottom));
			} else {
				border.setBottom(null);
			}
			if (top != null) {
				border.setTop(Boolean.valueOf(top));
			} else {
				border.setTop(null);
			}
			if (left != null) {
				border.setLeft(Boolean.valueOf(left));
			} else {
				border.setLeft(null);
			}
			if (right != null) {
				border.setRight(Boolean.valueOf(right));
			} else {
				border.setRight(null);
			}
			if (all != null) {
				border.setAll(Boolean.valueOf(all));
			} else {
				border.setAll(null);
			}
			if (outer != null) {
				border.setOuter(Boolean.valueOf(outer));
			} else {
				border.setOuter(null);
			}
			if (none != null) {
				border.setNone(Boolean.valueOf(none));
			} else {
				border.setNone(null);
			}
			
			operProp.setContent(content);
			operProp.setBorder(border);
			operProp.setContent(content);
			glx.setOperProp(operProp);
			glxList.add(glx);
		}
	}

	private int getColLeft(List<Glx> glxList, int i) {
		if (i == 0) {
			return 0;
		}
		Glx glx = glxList.get(i - 1);
		int tempWidth = glx.getWidth();
		if(glx.isHidden()){
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
//		if(gly.isHidden()){
//			tempHeight = -1;
//		}
		return gly.getTop() + tempHeight + 1;
	}

	private void bookToOlExcelCellList(int rowBeginIndex, int rowEndIndex,
			List<ExcelRow> rowList, List<Gly> glyList, List<Glx> glxList,
			ExcelSheet excelSheet, List<OneCell> newCellList) {
		getCellList(rowBeginIndex, rowEndIndex, rowList, glyList, glxList, excelSheet, newCellList);
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
				for (int j = 0; j <= cellList.size()-1; j++) {
					ExcelCell excelCell = cellList.get(j);
					if (excelCell != null) {
						OneCell cell = new OneCell();
						String highLight = excelCell.getExps().get(Constant.HIGHLIGHT);
						if ("true".equals(highLight)) {
							cell.setHighlight(true);
						} else {
							cell.setHighlight(false);
						}
						setCellProperty(excelCell, cell, i, j, glyList,glxList, excelSheet);
						if (excelSheet.getCols().get(j).isColumnhidden() || excelSheet.getRows().get(i).isRowhidden()) {
							int colspan = excelCell.getColspan();
							int rowspan = excelCell.getRowspan();
							if (colspan > 1 || rowspan > 1) {
								int[] mfCell = excelSheet.getMergFirstCell(i, j);
//								int tempY = mfCell[0];
								int tempX = mfCell[1];
								boolean flag = true;
//								for (int k = tempY; k < tempY + rowspan; k++) {
//									if (!excelSheet.getRows().get(k).isRowhidden()) {
//										flag = false;
//									}
//								}
								for (int k = tempX; k < tempX + colspan; k++) {
									if (!excelSheet.getCols().get(k).isColumnhidden()) {
										flag = false;
									}
								}
								if (flag) {
									cell.setHidden(true);
								}
							}else{
								cell.setHidden(true);
							}
						}
						newCellList.add(cell);
					}
				}
				Map<String, String> colMap = excelRow.getExps();
				Content content = glyList.get(i).getOperProp().getContent();
				String alignCol = colMap.get("align_vertical");
				content.setAlignCol(alignCol);
				String alignRow = colMap.get("align_level");
				content.setAlignRow(alignRow);
				String bd = colMap.get("font_weight");
				if (bd != null) {
					content.setBd(Boolean.valueOf(bd));
				}else{
					content.setBd(null);
				}
				String color = colMap.get("font_color");
				content.setColor(color);
				content.setFamily(colMap.get("font_family"));
				String italic = colMap.get("font_italic");
				if (italic != null) {
					content.setItalic(Boolean.valueOf(italic));
				} else {
					content.setItalic(null);
				}
				content.setSize(colMap.get("font_size"));
				content.setRgbColor(null);
				content.setTexts(null);
				CustomProp customProp = glyList.get(i).getOperProp().getCustomProp();
				Format format = glyList.get(i).getOperProp().getFormate();
				customProp.setBackground(colMap.get("fill_bgcolor"));
				format.setType(colMap.get("format"));
				format.setCurrencySign(colMap.get("currency"));
				format.setDateFormat(colMap.get("dateFormat"));
				String decimalPoint = colMap.get("decimalPoint");
				if (decimalPoint != null) {
					format.setDecimal(Integer.valueOf(decimalPoint));
				} else {
					format.setDecimal(null);
				}
				String thousandPoint = colMap.get("thousandPoint");
				if (thousandPoint != null) {
					format.setThousands(Boolean.valueOf(thousandPoint));
				} else {
					format.setThousands(null);
				}
				
				customProp.setComment(colMap.get("comment"));
				Border border = glyList.get(i).getOperProp().getBorder();
				String bottom = colMap.get("bottom");
				String top = colMap.get("top");
				String left = colMap.get("left");
				String right = colMap.get("right");
				String all = colMap.get("all");
				String outer = colMap.get("outer");
				String none = colMap.get("none");
				if(bottom != null){
					border.setBottom(Boolean.valueOf(bottom));
				}else{
					border.setBottom(null);
				}
				if(top != null){
					border.setTop(Boolean.valueOf(top));
				}else{
					border.setTop(null);
				}
				if(left != null){
					border.setLeft(Boolean.valueOf(left));
				}else{
					border.setLeft(null);
				}
				if(right != null){
					border.setRight(Boolean.valueOf(right));
				}else{
					border.setRight(null);
				}
				if(all != null){
					border.setAll(Boolean.valueOf(all));
				}else{
					border.setAll(null);
				}
				if(outer != null){
					border.setOuter(Boolean.valueOf(outer));
				}else{
					border.setOuter(null);
				}
				if(none != null){
					border.setNone(Boolean.valueOf(none));
				}else{
					border.setNone(null);
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
			if (topBorder > 0) {
				border.setTop(true);
			}
		}
		Excelborder excelBottomborder  = excelCellStyle.getBottomborder();
		if(excelBottomborder != null){
			short bottomBorder = excelBottomborder.getSort();
			if (bottomBorder >0) {
				border.setBottom(true);
			}
		}
		Excelborder excelLeftborder = excelCellStyle.getLeftborder();
		if(excelLeftborder != null){
			short leftBorder = excelLeftborder.getSort();
			if (leftBorder > 0) {
				border.setLeft(true);
			}
		}
		Excelborder excelRightborder = excelCellStyle.getRightborder();
		if(excelRightborder != null){
			short rightBorder = excelRightborder.getSort();
			if (rightBorder > 0) {
				border.setRight(true);
			}
		}
		cell.setWordWrap(excelCellStyle.isWraptext());
		Content content = cell.getContent();
		switch (excelCellStyle.getAlign()) {
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
			//content.setAlignRow("left");
			break;
		}
		switch (excelCellStyle.getValign()) {
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
			content.setAlignCol("middle");
			break;
		}
		ExcelFont excelFont = excelCellStyle.getFont();
		content.setFamily(excelFont.getFontname());
		if (700 == excelFont.getBoldweight()) {
			content.setBd(true);
		} else {
			content.setBd(false);
		}
		content.setItalic(excelFont.isItalic());
//		ExcelColor fontColor = excelFont.getColor();
//		if (fontColor != null) {
//			int[] rgb = fontColor.getRGBInt();
//			content.setColor(ExcelUtil.getRGB(rgb));
//			content.setColor("rgb(0,0,0)");
//		}
		content.setSize(excelFont.getSize() / 20 + "");
		content.setTexts(excelCell.getText());
		CustomProp customProp = cell.getCustomProp();
		Format formate = cell.getFormat();
		//if ("true".equals(excelSheet.getExps().get("ifUpload"))) {
		//if (excelCell.getExps().get("format") == null || "normal".equals(excelCell.getExps().get("format")) ) {
		//if (("".equals(excelCell.getShowText())|| excelCell.getShowText() == null) && !"".equals(excelCell.getText())) {
			CellFormateUtil.setShowText(excelCell, content,formate);
//			System.out.println(excelCell.getText());
//			String displayText = ExcelFormat.getShowText(excelCell);
//			content.setDisplayTexts(displayText);
//		} else {
//			String format = excelCell.getExps().get("format");
//			formate.setType(format);
//			formate.setThousands(Boolean.valueOf(excelCell.getExps().get("thousandPoint")));
//			if ("date".equals(format)) {
//				formate.setDateFormat(excelCell.getExps().get("dataFormate"));
//			}
//			String decimalPoint = excelCell.getExps().get("decimalPoint");
//			if (!StringUtil.isEmpty(decimalPoint)) {
//				formate.setDecimal(Integer.valueOf(decimalPoint));
//			}
//			formate.setCurrencySign(excelCell.getExps().get("currencySymbol"));
//			//excelCell.getShowText();
//			String showText = excelCell.getExps().get("displayText");
//			if(showText == null){
//				content.setDisplayTexts("");
//			}else{
//				content.setDisplayTexts(showText);
//			}
//		}
		
		//if(!StringUtil.isEmpty(excelCell.getMemo())){
			customProp.setComment(excelCell.getMemo());
		//}
		
		
		ExcelColor fontColor = excelFont.getColor();
		if (fontColor != null) {
			int[] rgb = fontColor.getRGBInt();
			content.setColor(ExcelUtil.getRGB(rgb));
			//content.setColor("rgb(0,0,0)");
		}
		// ExcelColor bgColor = excelCellStyle.getBgcolor();
		ExcelColor fgColor = excelCellStyle.getFgcolor();
		// if(bgColor != null){
		// int[] rgb = bgColor.getRGBInt();
		// customProp.setBackground(ExcelUtil.getRGB(rgb));
		// }else if(bgColor == null && fgColor != null){
		if (fgColor != null) {
			int[] rgb = fgColor.getRGBInt();
			customProp.setBackground(ExcelUtil.getRGB(rgb));
		}
		Occupy occupy = cell.getOccupy();
		int rowspan = excelCell.getRowspan();
		int colspan = excelCell.getColspan();
		int[] firstMergeCell = excelSheet.getMergFirstCell(i, j);
		if (firstMergeCell == null) {
			occupy.getX().add(glxList.get(j).getAliasX());
			occupy.getY().add(glyList.get(i).getAliasY());
			occupy.setRow(Integer.valueOf(glyList.get(i).getAliasY())-1);
			occupy.setCol(Integer.valueOf(glxList.get(j).getAliasX())-1);
		} else {
			int firstRow = firstMergeCell[0];
			int firstCol = firstMergeCell[1];
			occupy.setRow(firstRow);
			occupy.setCol(firstCol);
			for (int m = firstRow; m < firstRow + rowspan; m++) {
				occupy.getY().add(glyList.get(m).getAliasY());
			}
			for (int m = firstCol; m < firstCol + colspan; m++) {
//				System.out.println("=============="+i);
//				System.out.println("=============="+j);
				occupy.getX().add(glxList.get(m).getAliasX());
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

	public SpreadSheet positionExcel(ExcelSheet excelSheet, SpreadSheet spreadSheet, int height,ReturnParam returnParam,int end) {
		spreadSheet.setName(excelSheet.getName());
		List<Gly> glyList = spreadSheet.getSheet().getGlY();
		List<Glx> glxList = spreadSheet.getSheet().getGlX();
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
		returnParam.setMaxPixel(maxTop);
		returnParam.setMaxRowPixel(maxTop);
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
		returnParam.setMaxColPixel(maxLeft);
		// int oldrowBeginIndex = BinarySearch.rowsBinarySearch(startPixel,
		// glyList, 0, glyList.size()-1);
		int rowBeginIndex = 0;
		// int oldrowEndIndex = BinarySearch.rowsBinarySearch(endPixel, glyList,
		// 0, glyList.size()-1);
		int rowEndIndex = end;
		
		List<OneCell> newCellList = new ArrayList<OneCell>();
		spreadSheet.getSheet().setCells(newCellList);
		bookToOlExcelCellList(rowBeginIndex, rowEndIndex,rowList, glyList,glxList, excelSheet, newCellList);
		spreadSheet.getSheet().setGlY(glyList.subList(rowBeginIndex, rowEndIndex + 1));
		returnParam.setDataRowStartIndex(rowBeginIndex);
		returnParam.setDisplayRowStartAlias("1");
		returnParam.setDisplayColStartAlias("1");
		ExcelSheetFreeze excelfrozen = excelSheet.getFreeze();
		if (excelfrozen != null) {
			Frozen frozen = spreadSheet.getSheet().getFrozen();
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
//			frozen.setViewCol(excelfrozen.getFirstcol());
//			frozen.setViewRow(excelfrozen.getFirstrow());
		}
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
