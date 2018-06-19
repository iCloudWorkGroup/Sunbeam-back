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

import acmr.excel.ExcelHelper;
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
		baseDao.insert(excelId, mbook);//存储MBook对象
		List<ExcelSheet> excelSheets = book.getSheets();
		ExcelSheet excelSheet = excelSheets.get(0);
		MSheet msheet = new MSheet();
		String sheetId = excelId+0;
		msheet.setId(sheetId);
		msheet.setStep(0);
		msheet.setMaxrow(excelSheet.getMaxrow());
		msheet.setMaxcol(excelSheet.getMaxcol());
		msheet.setSheetName(excelSheet.getName());
		ExcelSheetFreeze freeze = excelSheet.getFreeze();
		if(null != freeze){
			msheet.setFreeze(freeze.isFreeze());
			msheet.setViewRowAlias(Integer.toString(freeze.getFirstrow()+1-freeze.getRow()));
			msheet.setViewColAlias(Integer.toString(freeze.getFirstcol()+1-freeze.getCol()));
			msheet.setRowAlias(Integer.toString(freeze.getFirstrow()+1));
			msheet.setColAlias(Integer.toString(freeze.getFirstcol()+1));
		}else{
			msheet.setFreeze(false);
			msheet.setViewColAlias("1");
			msheet.setViewRowAlias("1");
		}
		baseDao.insert(excelId, msheet);//存储MSheet对象
		
		
		List<ExcelColumn> cols = excelSheet.getCols();//得到所有的列
		List<MCol> mcols = new ArrayList<MCol>();//存贮列样式
		MRowColList cList = new MRowColList();//用于存贮简化的列
		cList.setId("cList");
		cList.setSheetId(sheetId);
		getMCol(cols,mcols,cList,sheetId);
		//mongoTemplate.insert(mcols, excelId);
		baseDao.insert(excelId, mcols);//存列表样式
		baseDao.insert( excelId,cList);//向数据库存贮简化的列信息
	
		List<ExcelRow> rows = excelSheet.getRows();
		List<Object> tempList = new ArrayList<Object>();//存储mcell对象及关系表对象
		List<MRow> mrows = new ArrayList<MRow>();
		MRowColList rList = new MRowColList();
		rList.setId("rList");
		rList.setSheetId(sheetId);
		getMRow(rows,mrows,rList, tempList,sheetId);
		baseDao.insert(excelId, rList);//向数据库存贮行信息
		//mongoTemplate.insert(mrows,excelId);
		baseDao.insert(excelId, mrows);//存储行样式
		
		List<Sheet> sheets = book.getNativeSheet();//找出合并的单元格
		getMergeCell(sheets,tempList,sheetId);
		
		long start = System.currentTimeMillis();
		//mongoTemplate.insert(tempList, excelId);
		baseDao.insert(excelId, tempList);
		long end = System.currentTimeMillis() - start;
		System.out.println("存储时间为:" + end);
		
		return true;
	
	}
	
	private void getMCol(List<ExcelColumn> cols,List<MCol> mcols,MRowColList cList,String sheetId){
		
		for (int i = 0; i < cols.size(); i++) {
			ExcelColumn excelColumn = cols.get(i);
			ExcelCellStyle style = excelColumn.getCellstyle();
			ExcelFont font = style.getFont();
			MCol mcol = new MCol();
			RowCol rc = new RowCol();
			
			rc.setAlias(excelColumn.getCode());
			if(i>0){
				
				rc.setPreAlias(cols.get(i-1).getCode());//存贮它前一行的别名
			}
			mcol.setSheetId(sheetId);
			mcol.setAlias(excelColumn.getCode());
			boolean hidden = excelColumn.isColumnhidden();
			if(hidden){
				mcol.setHidden(true);
				mcol.setWidth(excelColumn.getWidth());
				rc.setLength(0);
			}else{
				mcol.setWidth(excelColumn.getWidth());
				mcol.setHidden(false);
				rc.setLength(excelColumn.getWidth());
			}
			
			cList.getRcList().add(rc);//存简化列属性
			
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
			content.setExpress(style.getDataformat());
			
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
			
		}
	}
	
	
	@SuppressWarnings("unused")
	private void getMRow(List<ExcelRow> rows,List<MRow> mrows,MRowColList rList,List<Object> tempList,String sheetId){
		for (int i = 0; i < rows.size(); i++) {
			ExcelRow excelRow = rows.get(i);
			if (excelRow != null) {
				
				ExcelCellStyle style = excelRow.getCellstyle();
				ExcelFont   font = style.getFont();
				MRow mrow = new MRow();
				RowCol rc = new RowCol();
				
				rc.setAlias(excelRow.getCode());
				if(i>0){
					
					rc.setPreAlias(rows.get(i-1).getCode());//存贮它前一行的别名
				}
				
				mrow.setSheetId(sheetId);
				mrow.setAlias(excelRow.getCode());
				boolean hidden = excelRow.isRowhidden();
				if(hidden){
					mrow.setHidden(true);
					mrow.setHeight(excelRow.getHeight());
					rc.setLength(0);
				}else{
					mrow.setHeight(excelRow.getHeight());
					mrow.setHidden(false);
					rc.setLength(excelRow.getHeight());
				}
				
				rList.getRcList().add(rc);//存简化行属性
				
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
				content.setExpress(style.getDataformat());
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
				
				//存储没有合并的单元格属性
				List<ExcelCell> cells = excelRow.getCells();
					for (int j = 0; j < cells.size(); j++) {
						ExcelCell cell = cells.get(j);
					  if(cell!=null && cell.getColspan()<2 && cell.getRowspan()<2){//判断是否合并单元格
						int row = i+1;
						int col = j+1;
						MCell mcell = getMCell(cell,sheetId,row,col);
						tempList.add(mcell);
						//关系映射表
						MRowColCell mrcl = new MRowColCell();
						mrcl.setSheetId(sheetId);
						mrcl.setRow(row+"");
						mrcl.setCol(col+"");
						mrcl.setCellId(row + "_" + col);
						tempList.add(mrcl);
		                }
					}
					
					
			}
		}
		
	}
	
    private MCell getMCell(ExcelCell excelCell,String sheetId,int row,int col){
		ExcelCellStyle style = excelCell.getCellstyle();
		ExcelFont font = style.getFont();
		MCell mcell = new MCell();
		mcell.setSheetId(sheetId);
		mcell.setId(row+"_"+col);
		mcell.setRowspan(excelCell.getRowspan());
		mcell.setColspan(excelCell.getColspan());
		
		Border border = mcell.getBorder();
		Excelborder excelTopborder = style.getTopborder();
		if(excelTopborder != null){
			short topBorder = excelTopborder.getSort();
			if (topBorder == 2) {
				border.setTop(2);
			}else if(topBorder == 1){
				border.setTop(1);
			}else{
				border.setTop(0);
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
				border.setBottom(0);
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
				border.setLeft(0);
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
		
		content.setUnderline((int)font.getUnderline());
		content.setWordWrap(style.isWraptext());
		content.setType(excelCell.getType().name());
		content.setExpress(style.getDataformat());
		ExcelColor color = style.getFgcolor();
		if (color != null) {
			int[] rgb = color.getRGBInt();
			content.setBackground(ExcelUtil.getRGB(rgb));
		}else{
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
	 * @param sheets
	 * @param tempList
	 * @param relation
	 */
	private void getMergeCell(List<Sheet> sheets,List<Object> tempList,String sheetId){
		
		for(int i=0;i<sheets.size();i++){
			Sheet sh = sheets.get(i);
			int mergeCount = sh.getNumMergedRegions();
			for(int j=0;j<mergeCount;j++){
				
				CellRangeAddress ca = sh.getMergedRegion(j);
				int firstRow = ca.getFirstRow();
				int firstCol = ca.getFirstColumn();
				int lastRow = ca.getLastRow();
				int lastCol = ca.getLastColumn();
				Row row = sh.getRow(firstRow);
				XSSFCell cell =(XSSFCell) row.getCell(firstCol);
				ExcelSheet es = new ExcelSheet();
				ExcelCell excell = es.getExcelCell(cell);
				MRowColCell mrcc = new MRowColCell();
				//合并单元格的最后一个cell，将右、下边框的属性赋值给合并单元格
				Row lrow = sh.getRow(lastRow);
				XSSFCell lastcell =(XSSFCell) lrow.getCell(lastCol);
				XSSFCellStyle cs = lastcell.getCellStyle();
				excell.getCellstyle().setRightborder(ExcelHelper.getExcelBorder(cs.getBorderRight(), cs.getRightBorderXSSFColor()));
		        excell.getCellstyle().setBottomborder(ExcelHelper.getExcelBorder(cs.getBorderBottom(), cs.getBottomBorderXSSFColor()));		
				
		        
				int rid = firstRow + 1;
				int cid = firstCol + 1;
				excell.setRowspan(lastRow-firstRow+1);
				excell.setColspan(lastCol-firstCol+1);
				MCell mcell = getMCell(excell, sheetId, rid, cid);
				tempList.add(mcell);
				//关系映射表
				for(int k=firstRow;k<lastRow+1;k++){
					for(int l=firstCol;l<lastCol+1;l++){
						//多个单元格指向一个单元对象
						MRowColCell mrcl = new MRowColCell();
						mrcl.setSheetId(sheetId);
						mrcl.setRow((k+1)+"");
						mrcl.setCol((l+1)+"");
						mrcl.setCellId(rid + "_" + cid);
						tempList.add(mrcl);
					}
				}
				
			}
		}
		
	}

	@Override
	public CompleteExcel reload(String excelId, String sheetId, Integer rowBegin, Integer rowEnd, Integer colBegin,
			Integer colEnd) {
		//将步骤重置为0
		msheetDao.updateStep(excelId, sheetId, 0);
		CompleteExcel excel = new CompleteExcel();
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		int rowBeginIndex = getIndexByPixel(sortRList,rowBegin);
		int rowEndIndex = getIndexByPixel(sortRList, rowEnd);
		int colBeginIndex = getIndexByPixel(sortCList, colBegin);
		int colEndIndex = getIndexByPixel(sortCList, colEnd);
		int top;
		int left;
		if(rowBeginIndex == 0){
			top =0;
		}else{
			RowCol rc  = sortRList.get(rowBeginIndex);
			top = rc.getTop();
		}
		if(colBeginIndex == 0){
			left = 0;
		}else{
			RowCol rc  = sortCList.get(colBeginIndex);
			left = rc.getTop();
		}
		List<String> rowList = new ArrayList<String>();
		Map<String,Integer> rMap = new HashMap<String,Integer>();
		List<String> colList = new ArrayList<String>();
		Map<String,Integer> cMap = new HashMap<String,Integer>();
		for(int i = rowBeginIndex;i<rowEndIndex+1;i++){
			rowList.add(sortRList.get(i).getAlias());
			rMap.put(sortRList.get(i).getAlias(), i);
		}
		for(int j = colBeginIndex;j<colEndIndex+1;j++){
			colList.add(sortCList.get(j).getAlias());
			cMap.put(sortCList.get(j).getAlias(), j);
		}
		
		//行样式
		List<MRow> rList = mrowDao.getMRowList(excelId, sheetId, rowList);
		Map<String,MRow> rowMap = new HashMap<String,MRow>();
		for(MRow mr:rList){
			rowMap.put(mr.getAlias(), mr);
		}
		List<Gly> yList = new ArrayList<Gly>();
		for(int i=0;i<rowList.size();i++){
			String alias = rowList.get(i);
			Gly gly = new Gly(rowMap.get(alias));
			gly.setSort(i+rowBeginIndex);
			gly.setTop(getRowTop(yList, i, top));
		    yList.add(gly);
		}
		
		//列样式
		List<MCol> cList = mcolDao.getMColList(excelId, sheetId, colList);
		Map<String,MCol> colMap = new HashMap<String,MCol>();
		for(MCol mc:cList){
			colMap.put(mc.getAlias(), mc);
		}
		List<Glx> xList = new ArrayList<Glx>();
		for(int i=0;i<colList.size();i++){
			String alias = colList.get(i);
			Glx glx = new Glx(colMap.get(alias));
			glx.setSort(i+colBeginIndex);
			glx.setLeft(getColLeft(xList, i, left));
		    xList.add(glx);
		}
		
		//关系表
		List<MRowColCell> relationList = mrowColCellDao.getMRowColCellList(excelId, sheetId, rowList, colList);
	    List<String> cellIdList = new ArrayList<String>();
		for(MRowColCell mrcc:relationList){
			cellIdList.add(mrcc.getCellId());
		}
		List<MCell> cellList = mcellDao.getMCellList(excelId, sheetId, cellIdList);
		List<OneCell> ceList = new ArrayList<OneCell>();
		for(MCell mc:cellList){
			OneCell  ce = new OneCell(mc);
			String[] ids = mc.getId().split("_");
			Occupy oc = ce.getOccupy();
			if(mc.getRowspan()==1&&mc.getColspan()==1){
				oc.getRow().add(ids[0]);
				oc.getCol().add(ids[1]);
			}else{
				int indexRow = rMap.get(ids[0]);
				for(int i=0;i<mc.getRowspan();i++){
					
					oc.getRow().add(sortRList.get(indexRow).getAlias());
					indexRow++;
				}
				int indexCol = cMap.get(ids[1]);
				for(int i=0;i<mc.getColspan();i++){
					
					oc.getCol().add(sortCList.get(indexCol).getAlias());
					indexCol++;
				}
				
			}
			ceList.add(ce);
		}
		
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		SheetElement sheet = new SheetElement(msheet);
		RowCol row = sortRList.get(sortRList.size()-1);
		sheet.setMaxRowPixel(row.getTop()+row.getLength()+1);
		RowCol col = sortCList.get(sortCList.size()-1);
		sheet.setMaxColPixel(col.getTop()+col.getLength()+1);
		sheet.setCells(ceList);
		sheet.setGridLineRow(yList);
		sheet.setGridLineCol(xList);
		excel.getSheets().add(sheet);
		excel.setName("new Book");
		
		
		return excel;
		
	}
	
   private int getIndexByPixel(List<RowCol> sortRclist,int pixel){
		
		RowCol rowColTop = sortRclist.get(sortRclist.size() - 1);
		int maxTop = rowColTop.getTop() + rowColTop.getLength();
		
		int endPixel = 0;
		if(pixel==0){
			endPixel = 0;
		}else if (maxTop < pixel) {
			endPixel = maxTop;
		} else {
			endPixel =  pixel;
		}
		int end = BinarySearch.binarySearch(sortRclist, endPixel);
		return end;
	}
   
   private int getColLeft(List<Glx> glxList, int i,Integer left) {
		if (i == 0) {
			return left;
		}
		Glx glx = glxList.get(i - 1);
		int tempWidth = glx.getWidth();
		if(null!=glx.getHidden()&&glx.getHidden()){
			tempWidth = 0;
		}
		return glx.getLeft() + tempWidth + 1;
	}

	private int getRowTop(List<Gly> glyList, int i,int top) {
		if (i == 0) {
			return top;
		}
		Gly gly = glyList.get(i - 1);
		int tempHeight = gly.getHeight();
		if(null!=gly.getHidden()&&gly.getHidden()){
			tempHeight = 0;
		}

		return gly.getTop() + tempHeight + 1;
	}
   

}