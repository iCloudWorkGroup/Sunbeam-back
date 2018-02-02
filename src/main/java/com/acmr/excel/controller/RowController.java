package com.acmr.excel.controller;

import java.io.IOException;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;


import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.ModelAndView;

import acmr.excel.pojo.ExcelBook;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowLine;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;

/**
 * 单元格操作
 * 
 * @author jinhr
 *
 */
 @Controller
 @RequestMapping("/row")
public class RowController extends BaseController {


	/**
	 * 插入行
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/plus")
	public void rows_insert(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		RowOperate rowOperate = getJsonDataParameter(req, RowOperate.class);
		this.assembleData(req, resp,rowOperate,OperatorConstant.rowsinsert);
	}

	/**
	 * 删除行
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/reduce")
	public void rows_delete(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		RowOperate rowOperate = getJsonDataParameter(req, RowOperate.class);
		this.assembleData(req, resp,rowOperate,OperatorConstant.rowsdelete);
	}


	/**
	 * 行隐藏
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/hide")
	public void rows_hide(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		RowOperate colOperate = getJsonDataParameter(req, RowOperate.class);
		this.assembleData(req, resp,colOperate,OperatorConstant.rowshide);
	}
	/**
	 * 高度调整
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/adjust")
	public void rows_height(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		RowHeight rowHeight = getJsonDataParameter(req, RowHeight.class);
		this.assembleData(req, resp,rowHeight,OperatorConstant.rowsheight);
	}

    /**
	 * 增加行，用于初始化时向下滚动
	 */
    @RequestMapping("/plus-batch")
	public void addrowline(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		RowLine rowLine = getJsonDataParameter(req, RowLine.class);
		this.assembleData(req, resp,rowLine,OperatorConstant.addRowLine);
	}
}
