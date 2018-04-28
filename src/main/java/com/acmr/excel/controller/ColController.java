package com.acmr.excel.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.AddLine;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;

/**
 * 单元格操作
 * 
 * @author jinhr
 *
 */
 @Controller
 @RequestMapping("/col")
public class ColController extends BaseController {



	/**
	 * 列增加
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/plus")
	public void cols_insert(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colsinsert);
	}

	/**
	 * 删除列
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/reduce")
	public void cols_delete(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colsdelete);
	}

	/**
	 * 宽度调整
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/adjust")
	public void cols_width(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		ColWidth colWidth = getJsonDataParameter(req, ColWidth.class);
		this.assembleData(req, resp,colWidth,OperatorConstant.colswidth);
	}
	/**
	 * 列隐藏
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/hide")
	public void cols_hide(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colshide);
	}
	/**
	 * 显示列
	 * 
	 * @throws Exception
	 */
    @RequestMapping("/show")
	public void cols_cancelhide(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		ColOperate colOperate = getJsonDataParameter(req, ColOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.colhideCancel);
	}

    /**
	 * 增加列，用于初始化时向右滚动
	 */
    @RequestMapping("/plus-batch")
	public void addcolline(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		AddLine rowLine = getJsonDataParameter(req, AddLine.class);
		this.assembleData(req, resp,rowLine,OperatorConstant.addRowLine);
	}
}
