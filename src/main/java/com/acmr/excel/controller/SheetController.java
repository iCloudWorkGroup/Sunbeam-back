package com.acmr.excel.controller;

import java.io.IOException;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;



import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelSheet;

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
import com.acmr.excel.service.ExcelService;
import com.acmr.excel.service.PasteService;
import com.acmr.excel.service.StoreService;
import com.acmr.excel.service.impl.MongodbServiceImpl;
import com.acmr.excel.util.AnsycDataReturn;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.StringUtil;

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
		String excelId = req.getHeader("excelId");
		String curStep = req.getHeader("step");
		JsonReturn data = new JsonReturn("");
		if(StringUtil.isEmpty(excelId) || StringUtil.isEmpty(curStep)){
			resp.setStatus(400);
			return;
		}
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		int memStep = (int)mongodbServiceImpl.get(null, null, null);
		int cStep = 0;
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		ExcelBook excelBook = (ExcelBook) mongodbServiceImpl.get(null, null, null);
		
		if (cStep == memStep) {
			if (excelBook != null) {
				ExcelSheet excelSheet = excelBook.getSheets().get(0);
				ReturnParam returnParam = new ReturnParam();
				CompleteExcel excel = new CompleteExcel();
				SpreadSheet spreadSheet = new SpreadSheet();
				excel.getSpreadSheet().add(spreadSheet);
				spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam);
				data.setReturncode(Constant.SUCCESS_CODE);
				data.setReturndata(excel);
				data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
				data.setMaxRowPixel(returnParam.getMaxRowPixel());
				mongodbServiceImpl.set(excelId,excelBook);
			} else {
				data.setReturncode(Constant.CACHE_INVALID_CODE);
				data.setReturndata(Constant.CACHE_INVALID_MSG);
			}
		}else{
			for (int i = 0; i < 100; i++) {
				int mStep = (int)mongodbServiceImpl.get(null, null, null);
				if(cStep == mStep){
					if (excelBook != null) {
						ExcelSheet excelSheet = excelBook.getSheets().get(0);
						ReturnParam returnParam = new ReturnParam();
						CompleteExcel excel = new CompleteExcel();
						SpreadSheet spreadSheet = new SpreadSheet();
						excel.getSpreadSheet().add(spreadSheet);
						spreadSheet = excelService.openExcel(spreadSheet, excelSheet,rowBegin, rowEnd,returnParam);
						data.setReturncode(Constant.SUCCESS_CODE);
						data.setReturndata(excel);
						data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
						data.setMaxRowPixel(returnParam.getMaxRowPixel());
						mongodbServiceImpl.set(excelId, excelBook);
					} else {
						data.setReturncode(Constant.CACHE_INVALID_CODE);
						data.setReturndata(Constant.CACHE_INVALID_MSG);
					}
				}else{
					try {
						Thread.sleep(100);
					} catch (InterruptedException e) {
						e.printStackTrace();
					}
				}
			}
		}
		if("".equals(data.getReturndata())){
			data.setReturncode(-1);
		}
		this.sendJson(resp, data);
	}
	
	
	
	
	
	
	
	
	
	
	
}
