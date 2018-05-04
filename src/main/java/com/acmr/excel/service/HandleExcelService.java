package com.acmr.excel.service;



import java.io.UnsupportedEncodingException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.springframework.stereotype.Service;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellStyle;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelFont;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;
import acmr.excel.pojo.Excelborder;
import acmr.util.ListHashMap;

import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColorSet;
import com.acmr.excel.model.ColorSetRet;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.complete.StrandY;
import com.acmr.excel.model.history.ChangeArea;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.model.mongo.MExcel;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelColumn;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.service.impl.MongodbServiceImpl;
import com.acmr.excel.util.BinarySearch;
import com.acmr.excel.util.CellFormateUtil;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.StringUtil;
import com.acmr.excel.util.UUIDUtil;

/**
 * 操作excelservice
 * 
 * @author caosl
 */
@Service
public class HandleExcelService {
	/**
	 * 单元格类型
	 * 
	 * @author jinhr
	 *
	 */
	public enum CellUpdateType {
		align_level, align_vertical, frame, font_size, date_format, text, font_color, fill_bgcolor, font_italic, font_weight, font_family, word_wrap;
	}

	/**
	 * 行，列
	 * 
	 * @author jinhr
	 *
	 */

	public enum RC {
		row, col;
	}

	/**
	 * 获取excel唯一性id
	 * 
	 * @return id
	 */
	public String getExcelId() {
		return UUIDUtil.getUUID();
	}

	/**
	 * 创建一个默认的excel
	 * 
	 * @param excelId
	 *            excelId
	 */
	public void createNewExcel(String excelId,MongodbServiceImpl storeService) {
		MExcel mExcel = new MExcel();
		mExcel.setId(excelId);
		mExcel.setStep(0);
		mExcel.setMaxrow(100);
		mExcel.setMaxcol(26);
		storeService.set(excelId, mExcel);
		ExcelSheet sheet = new ExcelSheet();
		for (int i = 1; i < 27; i++) {
			ExcelColumn column = sheet.addColumn();
			column.setWidth(69);
			MExcelColumn mc = new MExcelColumn();
			mc.setExcelColumn(column);
			mc.setColSort(i-1);
			storeService.set(excelId, mc);
		}
		for (int i = 1; i < 101; i++) {
			ExcelRow row = sheet.addRow();
			row.setHeight(19);
			MExcelRow mr = new MExcelRow();
			mr.setExcelRow(row);
			mr.setRowSort(i-1);
			storeService.set(excelId, mr);
		}
	}

	/**
	 * 获取excel中最大的sheet数
	 * 
	 * @param excel
	 *            excel对象
	 * @return 最大的sheet数量
	 */

	public int getMaxSheetOfOneExcel(CompleteExcel excel) {
		int max = -1;
		if (excel.getSheets() != null && excel.getSheets().size() > 0) {
			for (SheetElement oneSheet : excel.getSheets()) {
				if (oneSheet.getSort() > max) {
					max = oneSheet.getSort();
				}
			}
		}
		return max;
	}

	/**
	 * 创建单元格
	 * 
	 * @param excel
	 *            CompleteExcel对象
	 * @param sheetId
	 *            sheetId
	 */

	public void createCell(ExcelBook excel, int sheetId, String rowAlais,
			String colAlais) {
		ExcelSheet excelSheet = excel.getSheets().get(sheetId);
		List<ExcelRow> rowList = excelSheet.getRows();
		List<ExcelColumn> columnList = excelSheet.getCols();
		int rowIndex = BinarySearch.getRowIndexByRowAlais(rowList, rowAlais) + 1;
		int colIndex = BinarySearch.getColIndexByColAlais(columnList, colAlais) + 1;
		ExcelCell excelCell = new ExcelCell();
		rowList.get(rowIndex).getCells().set(colIndex, excelCell);
	}

	/**
	 * 获取区域list
	 * 
	 * @param fitX
	 *            Glx集合
	 * @param fitY
	 *            Gly集合
	 * @param sheetElement
	 *            SheetElement对象
	 * @param startX
	 *            x轴开始索引
	 * @param startY
	 *            y轴开始索引
	 * @param endX
	 *            x轴结束索引
	 * @param endY
	 *            y轴结束索引
	 */
	public void getAreaXYList(List<Glx> fitX, List<Gly> fitY,
			SheetElement sheetElement, String startX, String startY,
			String endX, String endY) {
		// 查找受影响的相关行列
		List<Glx> xlist = sheetElement.getGridLineCol();
		List<Gly> ylist = sheetElement.getGridLineRow();
		int startx = Integer.valueOf(startX);
		int starty = Integer.valueOf(startY);
		int endx = startx;
		int endy = starty;
		if (!StringUtil.isEmpty(endX)) {
			endx = Integer.valueOf(endX);
		}
		if (!StringUtil.isEmpty(endY)) {
			endy = Integer.valueOf(endY);
		}
		fitX.addAll(xlist.subList(startx, endx + 1));
		fitY.addAll(ylist.subList(starty, endy + 1));
	}

	/**
	 * 获取区域中存在的格子
	 * 
	 * @param startX
	 *            x轴开始索引
	 * @param startY
	 *            y轴开始索引
	 * @param endX
	 *            x轴结束索引
	 * @param endY
	 *            y轴结束索引
	 * @return 存在的单元格的集合
	 */

	public List<Integer> getExistsCellInThisArea(CompleteExcel excel,
			int sheetId, String startX, String startY, String endX, String endY) {
		// 1,判断是否存在这个cell,将cell的索引放入list
		List<Integer> existsCell = new ArrayList<Integer>();
	/*	// 获取相应sheet
		SpreadSheet sheet = excel.getSpreadSheet().get(sheetId - 1);
		SheetElement sheetElement = sheet.getSheet();
		// 查询符合坐标范围的单元格区域
		List<Glx> fitX = new ArrayList<>();
		List<Gly> fitY = new ArrayList<>();
		getAreaXYList(fitX, fitY, sheetElement, startX, startY, endX, endY);
		// 查找cell
		
		// 2,根据y轴查询
		StrandY posiY = sheetElement.getPosi().getStrandY();
		Map<String, Map<String, Integer>> posiMapY = posiY.getAliasY();
		for (Gly fitYIndex : fitY) {
			Map<String, Integer> psoiX = posiMapY.get(fitYIndex.getAliasY());
			if (psoiX != null) {
				for (Glx xy : fitX) {
					if (psoiX.containsKey(xy.getAliasX())) {
						existsCell.add(psoiX.get(xy.getAliasX()));
					}
				}
			}

		}*/
		return existsCell;
	}

	/**
	 * 查询行或者列上存在的格子
	 * 
	 * @param rc
	 *            行列标记
	 * @param se
	 *            单元格属性类型
	 * @param index
	 *            索引
	 * @return
	 */
	public List<Integer> getExistsCellInThisRowOrCol(RC rc, SheetElement se,
			String index) {
		List<Integer> result = new ArrayList<>();
		/*Map<String, Integer> map = null;
		if (rc.equals(RC.col)) {
			map = se.getPosi().getStrandY().getAliasY().get(index);
		} else {
			map = se.getPosi().getStrandX().getAliasX().get(index);
		}
		if (map != null) {
			for (String key : map.keySet()) {
				result.add(map.get(key));
			}
		}*/
		return result;
	}

	private short getAlign(String style) {
		short retVal = -1;
		switch (style) {
		case "left":
			retVal = 1;
			break;
		case "center":
			retVal = 2;
			break;
		case "right":
			retVal = 3;
			break;
		default:
			break;
		}
		return retVal;
	}

	private short getVertical(String style) {
		short retVal = -1;
		switch (style) {
		case "top":
			retVal = 0;
			break;
		case "middle":
			retVal = 1;
			break;
		case "bottom":
			retVal = 2;
			break;
		default:
			break;
		}
		return retVal;
	}

	private void geBorder(String style, ExcelCellStyle excelCellStyle) {
		Excelborder excelBorder = new Excelborder();
		short value = 1;
		excelBorder.setSort(value);
		switch (style) {
		case "left":
			excelCellStyle.setLeftborder(excelBorder);
			break;
		case "right":
			excelCellStyle.setRightborder(excelBorder);
			break;
		case "top":
			excelCellStyle.setTopborder(excelBorder);
			break;
		case "bottom":
			excelCellStyle.setBottomborder(excelBorder);
			break;
		case "all":
			excelCellStyle.setLeftborder(excelBorder);
			excelCellStyle.setRightborder(excelBorder);
			excelCellStyle.setTopborder(excelBorder);
			excelCellStyle.setBottomborder(excelBorder);
			break;
		case "none":
			Excelborder newExcelBorder = new Excelborder();
			excelCellStyle.setLeftborder(newExcelBorder);
			excelCellStyle.setRightborder(newExcelBorder);
			excelCellStyle.setTopborder(newExcelBorder);
			excelCellStyle.setBottomborder(newExcelBorder);
			break;
		default:
			break;
		}
	}

	/**
	 * 设置文本内容
	 * 
	 * @param cell
	 * @param excelBook
	 */
	public void data(Cell cell,int step,MongodbServiceImpl mongodbServiceImpl,String excelId) {
//		Integer version = versionHistory.getVersion().get(step-1);
//		if(version == null){
//			version = 0;
//			
//		}
//		version += 1;
//		versionHistory.getVersion().put(step, version);
		int rowIndex = cell.getCoordinate().getStartRow();
		int colIndex = cell.getCoordinate().getStartCol();
		Map<String,Object> rowMap = new HashMap<String, Object>();
		rowMap.put("rowSort", rowIndex);
		MExcelRow mExcelRow = (MExcelRow)mongodbServiceImpl.get(rowMap,MExcelRow.class, excelId);
		Map<String,Object> colMap = new HashMap<String, Object>();
		colMap.put("colSort", colIndex);
		MExcelColumn mExcelColumn = (MExcelColumn)mongodbServiceImpl.get(colMap,MExcelColumn.class, excelId);
		Map<String,Object> cellMap = new HashMap<String, Object>();
		cellMap.put("rowId", mExcelRow.getExcelRow().getCode());
		cellMap.put("colId", mExcelColumn.getExcelColumn().getCode());
		MExcelCell mExcelCell = (MExcelCell)mongodbServiceImpl.get(cellMap,MExcelCell.class, excelId);
		if(mExcelCell == null){
			mExcelCell = new MExcelCell();
			ExcelCell excelCell = new ExcelCell();
			mExcelCell.setExcelCell(excelCell);
		}
		ExcelCell excelCell = mExcelCell.getExcelCell();
		String content = cell.getContent();
		try {
			content = java.net.URLDecoder.decode(content, "utf-8");
			excelCell.setType(CELLTYPE.STRING);
			excelCell.setText(content);
			excelCell.setValue(content);
			CellFormateUtil.autoRecognise(content, excelCell);
//			changeArea.setUpdateValue(excelCell);
//			history.getChangeAreaList().add(changeArea);
//			versionHistory.getMap().put(version, history);
			mExcelCell.setRowId(mExcelRow.getExcelRow().getCode());
			mExcelCell.setColId(mExcelColumn.getExcelColumn().getCode());
			mExcelCell.setId(mExcelRow.getExcelRow().getCode()+"_"+mExcelColumn.getExcelColumn().getCode());
			mongodbServiceImpl.updateCell(excelId, mExcelCell);
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		
//		History history = new History();
//		history.setOperatorType(OperatorConstant.textData);
//		ChangeArea changeArea = new ChangeArea();
//		changeArea.setColIndex(colIndex);
//		changeArea.setRowIndex(rowIndex);
//		if (excelCell == null) {
//			changeArea.setOriginalValue(excelCell);
//			excelCell = new ExcelCell();
//			excelCell.getCellstyle().setBgcolor(new ExcelColor(255, 255, 255));
//			cellList.set(colIndex, excelCell);
//		}else{
//			changeArea.setOriginalValue(excelCell.clone());
//		}
		
	}
	/**
	 * 设置文本内容
	 * 
	 * @param cell
	 * @param excelBook
	 */
	public void colorSet(Cell cell, ExcelBook excelBook) {
		int colStartIndex = cell.getCoordinate().getStartCol();
		int rowStartIndex = cell.getCoordinate().getStartRow();
		int colEndIndex = cell.getCoordinate().getEndCol();
		int rowEndIndex = cell.getCoordinate().getEndRow();
		ExcelSheet sheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		ListHashMap<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
		for (int i = rowList.size() - 1; i <= rowEndIndex; i++) {
			sheet.addRow();
		}
		for (int i = columnList.size() - 1; i <= colEndIndex; i++) {
			sheet.addColumn();
		}
		for(int i = rowStartIndex ; i <= rowEndIndex ;i++){
			ExcelRow excelRow = rowList.get(i);
			List<ExcelCell> cellList = excelRow.getCells();
			for(int j = colStartIndex;j<=colEndIndex;j++){
				ExcelCell excelCell = cellList.get(j);
				if (excelCell == null) {
					excelCell = new ExcelCell();
					cellList.set(j, excelCell);
				}
				ExcelCellStyle excelCellStyle = excelCell.getCellstyle();
				excelCellStyle.setBgcolor(ExcelUtil.getColor(cell.getBgcolor()));
				excelCellStyle.setFgcolor(ExcelUtil.getColor(cell.getBgcolor()));
				excelCellStyle.setPattern(Short.valueOf("1"));
			}
		}
	}
	
	public List<ColorSetRet> batchColorSet(ColorSet colorSet, ExcelBook excelBook) {
		List<Coordinate> coordinateList = colorSet.getCoordinates();
		ExcelSheet sheet = excelBook.getSheets().get(0);
		List<ColorSetRet> colorSetRetList = new ArrayList<ColorSetRet>();
		//int z = 0;
		for(Coordinate coordinate : coordinateList){
			//ColorSetRet cr = null;
			int colStartIndex = coordinate.getStartCol();
			int rowStartIndex = coordinate.getStartRow();
			int colEndIndex = coordinate.getEndCol();
			int rowEndIndex = coordinate.getEndRow();
			ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
			ListHashMap<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
			for (int i = rowList.size() - 1; i <= rowEndIndex; i++) {
				sheet.addRow();
			}
			for (int i = columnList.size() - 1; i < colEndIndex; i++) {
				sheet.addColumn();
			}
			for(int i = rowStartIndex ; i <= rowEndIndex ;i++){
				ExcelRow excelRow = rowList.get(i);
				List<ExcelCell> cellList = excelRow.getCells();
				for(int j = colStartIndex;j<=colEndIndex;j++){
					ExcelCell excelCell = cellList.get(j);
//					if(excelCell != null){
//						int colspan = excelCell.getColspan();
//						int rowspan = excelCell.getRowspan();
//						if(colspan != 1 || rowspan != 1){
//							int[] mergFirstCell = sheet.getMergFirstCell(i, j);
//							if(mergFirstCell[0] < rowStartIndex || mergFirstCell[0] > rowEndIndex || 
//									mergFirstCell[1] < colStartIndex || mergFirstCell[1] > colEndIndex){
//								cr = new ColorSetRet();
//								cr.setIndex(z);
//								cr.setErrorMessage("该区域存在合并的单元格，不能操作");
//								break;
//							}
//						}
//					}
					if (excelCell == null) {
						excelCell = new ExcelCell();
						cellList.set(j, excelCell);
					}
					ExcelCellStyle excelCellStyle = excelCell.getCellstyle();
					excelCellStyle.setBgcolor(ExcelUtil.getColor(colorSet.getColor()));
					excelCellStyle.setFgcolor(ExcelUtil.getColor(colorSet.getColor()));
					excelCellStyle.setPattern(Short.valueOf("1"));
				}
			}
		   //z++;
//		   if(cr != null){
//				colorSetRetList.add(cr);
//		   }
		}
		return colorSetRetList;
	}
	
	
	
	
	/**
	 * 更新单元格
	 * 
	 * @param type
	 *            单元格类型枚举值
	 * @param cell
	 *            单元格
	 */
	public void updateCells(CellUpdateType type, Cell cell, ExcelBook excelBook,VersionHistory versionHistory,int step) {
		ExcelSheet sheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)sheet.getRows();
		ListHashMap<ExcelColumn> columnList = (ListHashMap<ExcelColumn>)sheet.getCols();
		int rowBeginIndex = cell.getCoordinate().getStartRow();
		int colBeginIndex = cell.getCoordinate().getStartCol();
		int rowEndIndex = cell.getCoordinate().getEndRow();
		int colEndIndex = cell.getCoordinate().getEndCol();
		if (rowEndIndex == -1) {
			rowEndIndex = rowList.size() - 1;
			for (int i = rowBeginIndex; i <= rowEndIndex; i++) {
				Map<String, String> exps = rowList.get(i).getExps();
				setExps(exps, type, cell);
			}
		} 
		if (colEndIndex == -1 ) {
			colEndIndex = columnList.size() - 1;
			for (int i = colBeginIndex; i <= colEndIndex; i++) {
				Map<String, String> exps = columnList.get(i).getExps();
				setExps(exps, type, cell);
			}
		} 
		if (rowBeginIndex != -1 && rowEndIndex != -1 && colBeginIndex != -1&& colEndIndex != -1) {
			Integer version = versionHistory.getVersion().get(step-1);
			if(version == null){
				version = 0;
			}
			version += 1;
			versionHistory.getVersion().put(step, version);
			History history = new History();
			for (int i = rowBeginIndex; i <= rowEndIndex; i++) {
				ExcelRow excelRow = rowList.get(i);
				if (excelRow == null) {
					excelRow = new ExcelRow();
					rowList.set(i, excelRow);
				}
				List<ExcelCell> excelCellList = excelRow.getCells();
				for (int j = colBeginIndex; j <= colEndIndex; j++) {
					ExcelCell excelCell = excelCellList.get(j);
					ChangeArea changeArea = new ChangeArea();
					changeArea.setRowIndex(i);
					changeArea.setColIndex(j);
					if (excelCell == null) {
						changeArea.setOriginalValue(new ExcelCellStyle());
						excelCell = new ExcelCell();
						excelCellList.set(j, excelCell);
					}else{
						changeArea.setOriginalValue(excelCell.getCellstyle().clone());
					}
					int rowSpan = excelCell.getRowspan();
					int colSpan = excelCell.getColspan();
					if (rowSpan > 1 || colSpan > 1) {
						int[] cel = sheet.getMergFirstCell(i, j);
						if (i != cel[0] || j != cel[1]) {
							continue;
						}
					}
					ExcelCellStyle excelCellStyle = excelCell.getCellstyle();
					ExcelFont excelFont = excelCellStyle.getFont();
					if (type.equals(CellUpdateType.align_level)) { 
						history.setOperatorType(OperatorConstant.alignlevel);
						// 水平对齐
						String style = cell.getAlign();
						excelCellStyle.setAlign(getAlign(style));
					} else if (type.equals(CellUpdateType.align_vertical)) {
						history.setOperatorType(OperatorConstant.alignvertical);
						// 垂直对齐
						String style = cell.getAlign();
						excelCellStyle.setValign(getVertical(style));
					} else if (type.equals(CellUpdateType.frame)) {
						history.setOperatorType(OperatorConstant.fontfamily);
						String frameStyle = cell.getDirection();
						Excelborder excelBorder = new Excelborder();
						short value = 1;
						excelBorder.setSort(value);
						excelBorder.setColor(new ExcelColor(0, 0, 0));
						switch (frameStyle) {
						case "left":
							excelCellStyle.setLeftborder(excelBorder);
							break;
						case "right":
							excelCellStyle.setRightborder(excelBorder);
							break;
						case "top":
							excelCellStyle.setTopborder(excelBorder);
							break;
						case "bottom":
							excelCellStyle.setBottomborder(excelBorder);
							break;
						case "all":
							excelCellStyle.setLeftborder(excelBorder);
							excelCellStyle.setRightborder(excelBorder);
							excelCellStyle.setTopborder(excelBorder);
							excelCellStyle.setBottomborder(excelBorder);
							break;
						case "none":
							Excelborder newExcelBorder = new Excelborder();
							excelCellStyle.setLeftborder(newExcelBorder);
							excelCellStyle.setRightborder(newExcelBorder);
							excelCellStyle.setTopborder(newExcelBorder);
							excelCellStyle.setBottomborder(newExcelBorder);
							break;
						case "outer":
							Excelborder outerExcelBorder = new Excelborder();
							short val = 1;
							outerExcelBorder.setSort(val);
							outerExcelBorder.setColor(new ExcelColor());
							//if (colAlais.equals(startX) && rowAlais.equals(startY)) {
							if (j == colBeginIndex && i == rowBeginIndex) {
								// 左上
								excelCellStyle.setLeftborder(outerExcelBorder);
								excelCellStyle.setTopborder(outerExcelBorder);
							}
							//if (colAlais.equals(endX) && rowAlais.equals(endY)) {
							if (j == colEndIndex && i == rowEndIndex) {
								// 右下
								excelCellStyle.setRightborder(outerExcelBorder);
								excelCellStyle.setBottomborder(outerExcelBorder);
							}
							//if (colAlais.equals(endX) && rowAlais.equals(startY)) {
							if (j ==  colEndIndex && i == rowBeginIndex) {
								// 右上
								excelCellStyle.setRightborder(outerExcelBorder);
								excelCellStyle.setTopborder(outerExcelBorder);
							}
							//if (colAlais.equals(startX) && rowAlais.equals(endY)) {
							if (j == colBeginIndex && i == rowEndIndex) {
								// 左下
								excelCellStyle.setBottomborder(outerExcelBorder);
								excelCellStyle.setLeftborder(outerExcelBorder);
							}
							//if (!colAlais.equals(startX)&& !colAlais.equals(endX)&& rowAlais.equals(startY)) {
							if (j != colBeginIndex && j != colEndIndex && i == rowBeginIndex) {
								// 上中
								excelCellStyle.setTopborder(outerExcelBorder);
							}
							//if (!colAlais.equals(startX)&& !colAlais.equals(endX)&& colAlais.equals(endY)) {
							if (j != colBeginIndex && j != colEndIndex && i == rowEndIndex) {
								// 下中
								excelCellStyle.setBottomborder(outerExcelBorder);
							}
							//if (!rowAlais.equals(startY)&& !rowAlais.equals(endY)&& colAlais.equals(startX)) {
							if (i != rowBeginIndex && i != rowEndIndex && j == colBeginIndex) {
								// 左中
								excelCellStyle.setLeftborder(outerExcelBorder);
							}

							if (i != rowBeginIndex && i != rowEndIndex && j ==  colEndIndex ) {
								// 右中
								excelCellStyle.setRightborder(outerExcelBorder);
							}
							break;
						default:
							break;
						}
					} else if (type.equals(CellUpdateType.font_size)) {
						history.setOperatorType(OperatorConstant.fontsize);
						short size = Short.valueOf(cell.getSize());
						size *= 20;
						excelFont.setSize(size);
						excelCellStyle.setFont(excelFont);
					} else if (type.equals(CellUpdateType.font_family)) {
						history.setOperatorType(OperatorConstant.fontfamily);
						excelFont.setFontname(cell.getFamily());
						excelCellStyle.setFont(excelFont);
					} else if (type.equals(CellUpdateType.font_weight)) {
						history.setOperatorType(OperatorConstant.fontweight);
						if (cell.isWeight()) {
							short value = 700;
							excelFont.setBoldweight(value);
						} else {
							short value = 0;
							excelFont.setBoldweight(value);
						}
						excelCellStyle.setFont(excelFont);
					} else if (type.equals(CellUpdateType.font_italic)) {
						history.setOperatorType(OperatorConstant.fontitalic);
						excelFont.setItalic(Boolean.valueOf(cell.getItalic()));
						excelCellStyle.setFont(excelFont);
					} else if (type.equals(CellUpdateType.font_color)) {
						history.setOperatorType(OperatorConstant.fontcolor);
						excelFont.setColor(ExcelUtil.getColor(cell.getColor()));
						excelCellStyle.setFont(excelFont);
					} else if (type.equals(CellUpdateType.fill_bgcolor)) {
						history.setOperatorType(OperatorConstant.fillbgcolor);
						excelCellStyle.setBgcolor(ExcelUtil.getColor(cell.getColor()));
						excelCellStyle.setFgcolor(ExcelUtil.getColor(cell.getColor()));
						excelCellStyle.setPattern(Short.valueOf("1"));
					} else if (type.equals(CellUpdateType.word_wrap)) {
						history.setOperatorType(OperatorConstant.wordWrap);
						excelCellStyle.setWraptext(Boolean.valueOf(cell.getWordWrap()));
					}
					changeArea.setUpdateValue(excelCellStyle);
					history.getChangeAreaList().add(changeArea);
				}
			}
			versionHistory.getMap().put(version, history);
			
		}
	}
	/**
	 * 行列操作时设置属性
	 * @param exps
	 * @param type
	 * @param cell
	 */
	private void setExps(Map<String,String> exps,CellUpdateType type,Cell cell){
		String types = type.toString();
		switch (types) {
		case "font_weight":
			//粗体
			if("true".equals(cell.getIsBold())){
				exps.put(types, cell.getIsBold());
			}else{
				exps.remove(types);
			}
			break;
		case "font_italic":
			//斜体
			if("true".equals(cell.getItalic())){
				exps.put(types, cell.getItalic());
			}else{
				exps.remove(types);
			}
			break;
		case "frame":
			//边框
			if("none".equals(cell.getDirection())){
				exps.remove("bottom");
				exps.remove("top");
				exps.remove("left");
				exps.remove("right");
				exps.remove("all");
				exps.remove("outer");
			}else{
				exps.put(cell.getDirection(), cell.getDirection());
			}
			break;
		case "fill_bgcolor":
			//背景色
			if("rgb(255,255,255)".equals(cell.getBgcolor())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getBgcolor());
			}
			break;
		case "align_level":
			//水平位置
			if("left".equals(cell.getAlign())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getAlign());
			}
			break;
		case "align_vertical":
			//垂直位置
			if("middle".equals(cell.getAlign())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getAlign());
			}
			break;
		case "font_color":
			//字体颜色
			if("rgb(0,0,0)".equals(cell.getColor())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getColor());
			}
			break;
		case "font_family":
			//字体
			if("SimSun".equals(cell.getFamily())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getFamily());
			}
			break;
		case "font_size":
			//字号
			if("11".equals(cell.getSize())){
				exps.remove(types);
			}else{
				exps.put(types, cell.getSize());
			}
			break;
		default:
			break;
		}
	}
	
	
	
	
	/**
	 * 单元格格式
	 * @param type
	 * @param cellFormate
	 * @param excelBook
	 */
	public void setCellFormate(CellFormate cellFormate, ExcelBook excelBook,VersionHistory versionHistory,int step) {
		Integer version = versionHistory.getVersion().put(step, step);
		if(version == null){
			version = 0;
		}
		version += 1;
		History history = new History();
		history.setOperatorType(OperatorConstant.textDataformat);
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>)excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>)excelSheet.getCols();
		int startRowIndex = cellFormate.getCoordinate().getStartRow();
		int endRowIndex = cellFormate.getCoordinate().getEndRow();
		if (endRowIndex == -1) {
			endRowIndex = rowList.size() - 1;
		} 
		int startColIndex = cellFormate.getCoordinate().getStartCol();
		int endColIndex = cellFormate.getCoordinate().getEndCol();
		boolean colFlag = false;
		if (endColIndex == -1) {
			endColIndex = colList.size() - 1;
			colFlag = true;
		} 
		String formate = cellFormate.getFormat();
		for (int i = startRowIndex; i <= endRowIndex; i++) {
			if (colFlag) {
				Map<String, String> exps = rowList.get(i).getExps();
				exps.put("format", formate);
				exps.put("thousandPoint", cellFormate.isThousandPoint()+"");
				exps.put("currency", cellFormate.getCurrencySymbol());
				exps.put("dateFormat", cellFormate.getDateFormat());
				exps.put("decimalPoint", cellFormate.getDecimalPoint()+"");
			}
			List<ExcelCell> cellList = rowList.get(i).getCells();
			for (int j = startColIndex; j <= endColIndex; j++) {
				ChangeArea changeArea = new ChangeArea();
				changeArea.setColIndex(j);
				changeArea.setRowIndex(i);
				ExcelCell excelCell = cellList.get(j);
				if(excelCell == null){
					changeArea.setOriginalValue(null);
					excelCell = new ExcelCell();
					cellList.set(j, excelCell);
				}else{
					changeArea.setOriginalValue(excelCell.clone());
				}
				switch (formate) {
				case "normal":
					CellFormateUtil.setGeneral(excelCell);
					break;
				case "text":
					CellFormateUtil.setText(excelCell);
					break;
				case "number":
					CellFormateUtil.setNumber(excelCell, cellFormate.getDecimalPoint(), cellFormate.isThousandPoint());
					break;
				case "date":
					CellFormateUtil.setTime(excelCell,cellFormate.getDateFormat());
					break;
				case "currency":
					CellFormateUtil.setCurrency(excelCell, cellFormate.getDecimalPoint(), cellFormate.getCurrencySymbol());
					break;
				case "percent":
					CellFormateUtil.setPercent(excelCell, cellFormate.getDecimalPoint());
					break;
				default:
					break;
				}
				changeArea.setUpdateValue(excelCell);
				history.getChangeAreaList().add(changeArea);
				versionHistory.getMap().put(version,history);
				excelCell.getExps().put("format", formate);
			}
		}
		for (int j = startColIndex; j <= endColIndex; j++) {
			Map<String, String> colExps = colList.get(j).getExps();
			colExps.put("format", formate);
			colExps.put("thousandPoint", cellFormate.isThousandPoint()+"");
			colExps.put("currency", cellFormate.getCurrencySymbol());
			colExps.put("dateFormat", cellFormate.getDateFormat());
			colExps.put("decimalPoint", cellFormate.getDecimalPoint()+"");
		}
	}
	
	
	/**
	 * 设置文本注释
	 * 
	 * @param excelBook
	 * @param comment
	 */

	public void setComment(ExcelBook excelBook, Comment comment,VersionHistory versionHistory,int step) {
		Integer version = versionHistory.getVersion().get(step-1);
		if(version == null){
			version = 0;
			
		}
		version += 1;
		History history = new History();
		history.setOperatorType(OperatorConstant.commentset);
		versionHistory.getVersion().put(step, version);
		versionHistory.getMap().put(version, history);
		ExcelSheet excelSheet = excelBook.getSheets().get(0);
		ListHashMap<ExcelRow> rowList = (ListHashMap<ExcelRow>) excelSheet.getRows();
		ListHashMap<ExcelColumn> colList = (ListHashMap<ExcelColumn>) excelSheet.getCols();
		int startRowIndex = comment.getCoordinate().getStartRow();
		int endRowIndex = comment.getCoordinate().getEndRow();
		boolean rowFlag = false;
		if (endRowIndex == -1) {
			endRowIndex = rowList.size() - 1;
			rowFlag = true;
		} 
		int startColIndex = comment.getCoordinate().getStartCol();
		int endColIndex = comment.getCoordinate().getEndCol();
		boolean colFlag = false;
		if (endColIndex == -1) {
			endColIndex = colList.size() - 1;
			colFlag = true;
		} 
		String com = comment.getComment();
//		if(com == null){
//			com = "";
//		}
		for (int i = startRowIndex; i <= endRowIndex; i++) {
			if(rowFlag){
				Map<String, String> exps = rowList.get(i).getExps();
				exps.put("comment",com );
			}
			List<ExcelCell> cellList = rowList.get(i).getCells();
			for (int j = startColIndex; j <= endColIndex; j++) {
				ExcelCell excelCell = cellList.get(j);
				ChangeArea changeArea = new ChangeArea();
				changeArea.setColIndex(j);
				changeArea.setRowIndex(i);
				if(excelCell == null){
					changeArea.setOriginalValue(null);
					excelCell = new ExcelCell();
					cellList.set(j, excelCell);
				}else{
					changeArea.setOriginalValue(excelCell.getMemo());
				}
				excelCell.setMemo(com);
				changeArea.setUpdateValue(com);
				history.getChangeAreaList().add(changeArea);
			}
		}
		if(colFlag){
			for (int j = startColIndex; j <= endColIndex; j++) {
				Map<String, String> colExps = colList.get(j).getExps();
				colExps.put("comment", com);
			}
		}
		
	}
}
