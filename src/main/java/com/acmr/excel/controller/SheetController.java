package com.acmr.excel.controller;

import java.io.IOException;
import java.util.List;

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
import com.acmr.excel.model.complete.SpreadSheet;
import com.acmr.excel.model.copy.Copy;
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
		String excelId = req.getHeader("excelId");
		String curStep = req.getHeader("step");
		JsonReturn data = new JsonReturn("");
		if(StringUtil.isEmpty(excelId) || StringUtil.isEmpty(curStep)){
			resp.setStatus(400);
			return;
		}
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		int memStep = 0 ;
				//(int)mongodbServiceImpl.get(null, null, null);
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		int colBegin = openExcel.getLeft();
		int colEnd = openExcel.getRight() == 0 ? 2000 : openExcel.getRight() ;
//		rowBegin = mongodbServiceImpl.getIndexByPixel(excelId, rowBegin, "rList");
//		rowEnd = mongodbServiceImpl.getIndexByPixel(excelId, rowEnd, "rList");
//		colBegin = mongodbServiceImpl.getIndexByPixel(excelId, colBegin, "cList");
//		colEnd = mongodbServiceImpl.getIndexByPixel(excelId, colEnd, "cList");
		List<RowCol> rowList = mongodbServiceImpl.getRCList(excelId, "rList");
		List<RowCol> colList = mongodbServiceImpl.getRCList(excelId, "cList");
		rowBegin = mongodbServiceImpl.getIndex(rowList, rowBegin);
		rowEnd = mongodbServiceImpl.getIndex(rowList, rowEnd);
		colBegin = mongodbServiceImpl.getIndex(colList, colBegin);
		colEnd = mongodbServiceImpl.getIndex(colList, colEnd);
//		long ba = 0;
//		long ba2 = 0;
//		long ba3 = 0;
//		long ba4 = 0;
		if (cStep == memStep) {
//			ba = System.currentTimeMillis();
			ExcelSheet excelSheet = mongodbServiceImpl.getSheetBySort(rowBegin,rowEnd,colBegin,colEnd, excelId);
//			ba2 = System.currentTimeMillis();
			
			if (excelSheet != null) {
				ReturnParam returnParam = new ReturnParam();
				CompleteExcel excel = new CompleteExcel();
				SpreadSheet spreadSheet = new SpreadSheet();
				excel.getSpreadSheet().add(spreadSheet);
				spreadSheet = excelService.openExcel(spreadSheet, excelSheet, rowBegin, rowEnd, colBegin, colEnd, returnParam, rowList, colList);
				data.setReturncode(Constant.SUCCESS_CODE);
				data.setReturndata(excel);
				data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
				data.setMaxRowPixel(rowList.get(rowList.size()-1).getTop());
			} else {
				data.setReturncode(Constant.CACHE_INVALID_CODE);
				data.setReturndata(Constant.CACHE_INVALID_MSG);
			}
		}
//		System.out.println("getSheetBySort =========" +(ba2-ba) );
//		System.out.println("excelService.openExcel =========" +(ba4-ba3) );
//		else{
//			for (int i = 0; i < 100; i++) {
//				int mStep = (int)mongodbServiceImpl.get(null, null, null);
//				if(cStep == mStep){
//					if (excelBook != null) {
//						ExcelSheet excelSheet = excelBook.getSheets().get(0);
//						ReturnParam returnParam = new ReturnParam();
//						CompleteExcel excel = new CompleteExcel();
//						SpreadSheet spreadSheet = new SpreadSheet();
//						excel.getSpreadSheet().add(spreadSheet);
//						spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,colBegin,colEnd,returnParam);
//						data.setReturncode(Constant.SUCCESS_CODE);
//						data.setReturndata(excel);
//						data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
//						data.setMaxRowPixel(returnParam.getMaxRowPixel());
//						mongodbServiceImpl.set(excelId, excelBook);
//					} else {
//						data.setReturncode(Constant.CACHE_INVALID_CODE);
//						data.setReturndata(Constant.CACHE_INVALID_MSG);
//					}
//				}else{
//					try {
//						Thread.sleep(100);
//					} catch (InterruptedException e) {
//						e.printStackTrace();
//					}
//				}
//			}
//		}
		if("".equals(data.getReturndata())){
			data.setReturncode(-1);
		}
		long b2 = System.currentTimeMillis();
		System.out.println("openexcel=====================" + (b2 - b1));
//		System.out.println("========================================================");
		this.sendJson(resp, data);
	}
	
	
	
	
	
	
	
	
	
	
	
}
