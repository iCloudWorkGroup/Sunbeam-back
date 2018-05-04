package com.acmr.excel.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.Paste;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.ReturnParam;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.mongo.MExcelSheet;
import com.acmr.excel.model.position.OpenExcel;
import com.acmr.excel.model.position.RowCol;
import com.acmr.excel.service.ExcelService;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.impl.MongodbServiceImpl;
import com.acmr.excel.util.AnsycDataReturn;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.StringUtil;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;

/**
 * SHEET操作
 * 
 * @author jinhr
 *
 */
@Controller
@RequestMapping("/sheet")
public class SheetController extends BaseController {
	@Resource
	private MongodbServiceImpl mongodbServiceImpl;
	@Resource
	private PasteService pasteService; 
	 @Resource
	private ExcelService excelService;

	/**
	 * 新建sheet
	 * 
	 * @throws IOException
	 */
	public void create(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 修改sheet
	 * 
	 */
	public void update(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 删除sheet
	 * 
	 * @throws IOException
	 */
	public void delete(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
	}

	/**
	 * 冻结
	 * 
	 * @throws Exception
	 */
    @RequestMapping("/frozen")
	public void frozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.frozen);
	}

	/**
	 * 取消冻结
	 * 
	 * @throws Exception
	 */
    @RequestMapping("/unfrozen")
	public void unFrozen(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Frozen frozen = getJsonDataParameter(req, Frozen.class);
		this.assembleData(req, resp,frozen,OperatorConstant.unFrozen);
	}

	
	
	
	/**
	 * 回退
	 */
    @RequestMapping("/undo")
	public void undo(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		this.assembleData(req, resp,null,OperatorConstant.undo);
	}
	/**
	 * 前进
	 */
    @RequestMapping("/redo")
	public void redo(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		this.assembleData(req, resp,null,OperatorConstant.redo);
	}
	/**
	 * 外部粘贴
	 * @throws IOException
	 */
	@RequestMapping("/paste")
	public void paste(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Paste paste = getJsonDataParameter(req, Paste.class);
		ExcelBook excelBook = (ExcelBook)mongodbServiceImpl.get(null, null, null);
		boolean isAblePasteResult = pasteService.isAblePaste(paste, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, paste, OperatorConstant.paste);
			ansycDataReturn.setIsLegal(true);
			this.sendJson(resp, ansycDataReturn);
		}else{
			ansycDataReturn.setIsLegal(false);
			this.sendJson(resp, ansycDataReturn);
		}
		
	}
	/**
	 * 内部复制粘贴
	 * @throws IOException
	 */
	@RequestMapping("/copy")
	public void copy(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)mongodbServiceImpl.get(null, null, null);
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.copy);
			ansycDataReturn.setIsLegal(true);
			this.sendJson(resp, ansycDataReturn);
		}else{
			ansycDataReturn.setIsLegal(false);
			this.sendJson(resp, ansycDataReturn);
		}
		
	}
	/**
	 * 剪切粘贴
	 * @throws IOException
	 */
	@RequestMapping("/cut")
	public void cut(HttpServletRequest req,HttpServletResponse resp) throws Exception{
		String excelId = req.getHeader("excelId");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		ExcelBook excelBook = (ExcelBook)mongodbServiceImpl.get(null, null, null);
		boolean isAblePasteResult = pasteService.isCopyPaste(copy, excelBook);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, copy, OperatorConstant.cut);
			ansycDataReturn.setIsLegal(true);
			this.sendJson(resp, ansycDataReturn);
		}else{
			ansycDataReturn.setIsLegal(false);
			this.sendJson(resp, ansycDataReturn);
		}
	}
	
	
	/**
	 * 通过像素动态加载excel
	 */
	@RequestMapping("/area")
	public void openexcel(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		long b1 = System.currentTimeMillis();
		String excelId = req.getHeader("X-Book-Id");
		String curStep = req.getHeader("X-Step");
		
		if(StringUtil.isEmpty(excelId) || StringUtil.isEmpty(curStep)){
			resp.setStatus(400);
			return;
		}
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		int memStep = 0 ;
		
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		int colBegin = openExcel.getLeft();
		int colEnd = openExcel.getRight() == 0 ? 2000 : openExcel.getRight() ;
		
		List<RowCol> sortRcList = new ArrayList<RowCol>();//存贮整理顺序后的行
		List<RowCol> sortClList = new ArrayList<RowCol>();//存储整理顺序后的列
		Map<String,RowCol> newRcMap = new HashMap<String,RowCol>();
		Map<String,RowCol> newClMap = new HashMap<String,RowCol>();
		mongodbServiceImpl.getRowList(sortRcList, newRcMap, excelId);
		mongodbServiceImpl.getColList(sortClList, newClMap, excelId);
		int rowBeginIndex = mongodbServiceImpl.getIndexByPixel(sortRcList, rowBegin);
		int rowEndIndex = mongodbServiceImpl.getIndexByPixel(sortRcList,rowEnd);
		int colBeginIndex = mongodbServiceImpl.getIndexByPixel(sortClList,colBegin);
		int colEndIndex = mongodbServiceImpl.getIndexByPixel(sortClList,colEnd);
		CompleteExcel excel = new CompleteExcel();
		SheetElement sheet = new SheetElement();
		if (cStep == memStep) {

			ExcelSheet excelSheet = mongodbServiceImpl.getSheetBySort(rowBeginIndex,rowEndIndex,colBeginIndex,colEndIndex, excelId,sortRcList,sortClList);
			
			if (excelSheet != null) {
				
				
				
				excel.getSheets().add(sheet);
				sheet = excelService.openExcel(sheet, excelSheet, rowBeginIndex, rowEndIndex, colBeginIndex, colEndIndex,  sortRcList, sortClList);
			
			} else {
				
			}
		}

		
		long b2 = System.currentTimeMillis();
		System.out.println("openexcel=====================" + (b2 - b1));

		this.sendJson(resp, sheet);
	}
	
	
	
	
	
	
	
	
	
	
	
}
