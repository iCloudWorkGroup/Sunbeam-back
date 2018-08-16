package com.acmr.excel.controller;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OpenExcel;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.OuterPaste;
import com.acmr.excel.model.RowOrColExpand;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.model.complete.Glx;
import com.acmr.excel.model.complete.Gly;
import com.acmr.excel.model.complete.OneCell;
import com.acmr.excel.model.complete.SheetElement;
import com.acmr.excel.model.complete.Validate;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MBookService;
import com.acmr.excel.service.MSheetService;
import com.acmr.excel.util.AnsycDataReturn;
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
	private MSheetService msheetService;
	@Resource
	private MBookService mbookService;
	

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
		String excelId = req.getHeader("X-Book-Id");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		OuterPaste outerPaste = getJsonDataParameter(req, OuterPaste.class);
		
		boolean isAblePasteResult = msheetService.isAblePaste(outerPaste, excelId);
		AnsycDataReturn ansycDataReturn = new AnsycDataReturn();
		if(isAblePasteResult){
			this.assembleData(req, resp, outerPaste, OperatorConstant.paste);
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
		String excelId = req.getHeader("X-Book-Id");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		boolean isAblePasteResult = msheetService.isCopy(copy, excelId);
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
		String excelId = req.getHeader("X-Book-Id");
		if (StringUtil.isEmpty(excelId)) {
			resp.setStatus(400);
			return;
		}
		Copy copy = getJsonDataParameter(req, Copy.class);
		
		boolean isAblePasteResult = msheetService.isCut(copy, excelId);
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
		String sheetId = excelId+0;
		int memStep = 0 ;
		int cStep = 0;
		
		if(!StringUtil.isEmpty(curStep)){
			cStep = Integer.valueOf(curStep);
		}
		if(cStep>0){
		  while(true){
			memStep = msheetService.getStep(excelId,sheetId);
			if(cStep!=memStep){
				try {
					Thread.sleep(100);
				} catch (InterruptedException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}else{
				break;
			}
		  }
		}
       
		
		if(StringUtil.isEmpty(excelId) || StringUtil.isEmpty(curStep)){
			resp.setStatus(400);
			return;
		}
		OpenExcel openExcel = getJsonDataParameter(req, OpenExcel.class);
		int rowBegin = openExcel.getTop();
		int rowEnd = openExcel.getBottom();
		int colBegin = openExcel.getLeft();
		int colEnd = openExcel.getRight() == 0 ? 2000 : openExcel.getRight() ;
		
		CompleteExcel excel =  mbookService.reload(excelId, sheetId, rowBegin, rowEnd, colBegin, colEnd,1);
		SheetElement sheet = null;
		if(excel.getSheets().size()>0){
		   sheet = new SheetElement(excel.getSheets().get(0));
		   sheet.setFrozen(null);
		   sheet.setProtect(null);
		   sheet.setSort(null);
		}else{
			sheet = new SheetElement();
			sheet.setFrozen(null);
			sheet.setProtect(null);
			sheet.setSort(null);
			/*List<OneCell> cells = new ArrayList<OneCell>();
			List<Gly> gridLineRow = new ArrayList<Gly>();
			List<Glx> gridLineCol = new ArrayList<Glx>();
			List<Validate> validate = new ArrayList<Validate>();
			sheet.setCells(cells);
			sheet.setGridLineRow(gridLineRow);
			sheet.setGridLineCol(gridLineCol);
			sheet.setValidate(validate);*/
			
		}
		long b2 = System.currentTimeMillis();
		System.out.println("openexcel=====================" + (b2 - b1));

		this.sendJson(resp, sheet);
	}
	
	/**
	 * 通过像素动态加载excel
	 */
	@RequestMapping("/expand")
	public void addRowOrCol(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		RowOrColExpand expand = getJsonDataParameter(req, RowOrColExpand.class);
		this.assembleData(req, resp,expand,OperatorConstant.expand);
	}
	
	/**
	 * sheet保护
	 * @param req
	 * @param resp
	 * @throws IOException
	 */
	@RequestMapping("/protect")
	public void protect(HttpServletRequest req, HttpServletResponse resp) throws IOException {
		String excelId = req.getHeader("X-Book-Id");
		MSheet msheet = getJsonDataParameter(req, MSheet.class);
		msheetService.updateProtect(excelId, msheet);
		
	}
	
}
