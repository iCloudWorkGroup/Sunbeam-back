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
import com.acmr.excel.model.ColorSet;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.CellFormate.CellFormate;
import com.acmr.excel.model.comment.Comment;
import com.acmr.excel.model.complete.rows.ColOperate;
import com.acmr.excel.model.complete.rows.RowOperate;
import com.acmr.excel.service.HandleExcelService;

/**
 * 单元格操作
 * 
 * @author jinhr
 *
 */
 @Controller
 @RequestMapping("/cell")
public class CellController extends BaseController {

	/**
	 * 合并单元格
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/merge")
	public void merge(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.merge);
	}

	/**
	 * 单元格拆分
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/split")
	public void merge_delete(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.mergedelete);
	}

	/**
	 * 边框操作
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/border")
	public void frame(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.frame);
	}


	/**
	 * 水平对齐
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/align-landscape")
	public void align_level(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.alignlevel);
	}

	/**
	 * 垂直对齐
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/align-portrait")
	public void align_vertical(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp,cell,OperatorConstant.alignvertical);
	}

	/**
	 * 字号
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/font-size")
	public void font_size(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontsize);
	}

	/**
	 * 风格
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/font-family")
	public void font_family(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontfamily);
	}

	/**
	 * 粗细
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/font-weight")
	public void font_weight(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontweight);
	}

	/**
	 * 斜体
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/font-italic")
	public void font_italic(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontitalic);
	}

	/**
	 * 字体颜色
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/font-color")
	public void font_color(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fontcolor);
	}
	/**
	 * 自动换行
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */
	@RequestMapping("/wordwrap")
	public void wordwrap(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.wordWrap);
	}
	/**
	 * 背景颜色
	 * 
	 * @throws IOException
	 */
	@RequestMapping("/bg")
	public void fill_bgcolor(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		// 接收参数，定义返回
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.fillbgcolor);
	}

	/**
	 * 编辑单元格中数据内容
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/content")
	public void data(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Cell cell = getJsonDataParameter(req, Cell.class);
		this.assembleData(req, resp, cell, OperatorConstant.textData);
	}
	/**
	 * 设置内容数据类型
	 * 
	 * @throws IOException
	 */
    @RequestMapping("/format")
	public void data_format(HttpServletRequest req, HttpServletResponse resp)
			throws Exception {
		CellFormate cellFormate = getJsonDataParameter(req, CellFormate.class);
		this.assembleData(req, resp, cellFormate, OperatorConstant.textDataformat);
	}

	/**
	 * 批量设置备注
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */
    @RequestMapping("/comment-plus")
	public void comment_plus(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		Comment comment = getJsonDataParameter(req, Comment.class);
		this.assembleData(req, resp, comment, OperatorConstant.commentset);
	}
	
    @RequestMapping("/comment-reduce")
	public void comment_reduce(HttpServletRequest req, HttpServletResponse resp) throws Exception {
    	comment_plus(req, resp);
	}
	
	
//    @RequestMapping("/bg-batch")
//	public void color_set(HttpServletRequest req, HttpServletResponse resp) throws Exception {
//		Cell cell = getJsonDataParameter(req, Cell.class);
//		this.assembleData(req, resp, cell, OperatorConstant.colorset);
//	}
    @RequestMapping("/bg-batch")
	public void batchcolorset(HttpServletRequest req, HttpServletResponse resp) throws Exception {
    	ColorSet cell = getJsonDataParameter(req, ColorSet.class);
		this.assembleData(req, resp, cell, OperatorConstant.batchcolorset);
	}

}
