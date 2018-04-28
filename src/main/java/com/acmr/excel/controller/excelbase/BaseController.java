package com.acmr.excel.controller.excelbase;

import java.io.BufferedReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.InputStreamReader;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;


import org.apache.log4j.Logger;
import org.springframework.web.servlet.ModelAndView;
import org.springframework.web.servlet.mvc.multiaction.MultiActionController;

import com.acmr.excel.model.Constant;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.StringUtil;
import com.acmr.mq.Model;
import com.acmr.mq.producer.queue.QueueSender;
import com.alibaba.fastjson.JSON;


/**
 * excel base action
 * 
 * @author caosl
 */

public class BaseController extends MultiActionController {
	private static Logger logger = Logger.getLogger(BaseController.class);
	@Resource
	private QueueSender queueSender;

	
	/**
	 * session缓存
	 * 
	 * @param req
	 *            request对象
	 * @return session对象
	 */

	protected HttpSession getSession(HttpServletRequest req) {
		return req.getSession();
	}



	/**
	 * 把json串转化为实际对象
	 * 
	 * @param request
	 *            request对象
	 * @param t
	 *            泛型对象
	 * @return 实际对象
	 * @throws IOException
	 */

	protected <T> T getJsonDataParameter(HttpServletRequest request,
			Class<T> clazz) throws IOException {
		String body = null;
		StringBuilder stringBuilder = new StringBuilder();
		BufferedReader bufferedReader = null;
		try {
			InputStream inputStream = request.getInputStream();
			if (inputStream != null) {
				bufferedReader = new BufferedReader(new InputStreamReader(
						inputStream));
				char[] charBuffer = new char[128];
				int bytesRead = -1;
				while ((bytesRead = bufferedReader.read(charBuffer)) > 0) {
					stringBuilder.append(charBuffer, 0, bytesRead);
				}
			} else {
				stringBuilder.append("");
			}
		} catch (IOException ex) {
			throw ex;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (IOException ex) {
					throw ex;
				}
			}
		}
		body = stringBuilder.toString();
		if (StringUtil.isEmpty(body)) {
			try {
				return clazz.newInstance();
			} catch (InstantiationException | IllegalAccessException e) {
				e.printStackTrace();
			}
		}
		return JSON.parseObject(body, clazz);
	}

	protected String getBody(HttpServletRequest request) throws IOException {
		String body = null;
		StringBuilder stringBuilder = new StringBuilder();
		BufferedReader bufferedReader = null;
		try {
			InputStream inputStream = request.getInputStream();
			if (inputStream != null) {
				bufferedReader = new BufferedReader(new InputStreamReader(
						inputStream));
				char[] charBuffer = new char[128];
				int bytesRead = -1;
				while ((bytesRead = bufferedReader.read(charBuffer)) > 0) {
					stringBuilder.append(charBuffer, 0, bytesRead);
				}
			} else {
				stringBuilder.append("");
			}
		} catch (IOException ex) {
			throw ex;
		} finally {
			if (bufferedReader != null) {
				try {
					bufferedReader.close();
				} catch (IOException ex) {
					throw ex;
				}
			}
		}

		body = stringBuilder.toString();
		return body;
	}

	protected void sendJson(HttpServletResponse resp, Object data) {
		// resp.setHeader("Access-Control-Allow-Origin", "*");
		resp.setContentType("application/json; charset=utf-8");
			try {
				resp.getWriter().print(JSON.toJSONString(data));
			} catch (IOException e) {
				e.printStackTrace();
			}
	}
	protected void assembleData(HttpServletRequest req,HttpServletResponse resp,Object object,int reqPath){
		JsonReturn data = new JsonReturn("");
		String step = req.getHeader("step");
		String excelId = req.getHeader("excelId");
		if(StringUtil.isEmpty(step)){
			step = "1";
		}
		Model model = new Model();
		int index = Integer.valueOf(step);
		model.setStep(index);
		model.setReqPath(reqPath);
		model.setObject(object);
		model.setExcelId(excelId);
		logger.info("**********发送excelId:"+excelId+"====step:"+step+"===reqPath:"+reqPath);	
		queueSender.send(Constant.queueName, model);
		data.setReturncode(Constant.SUCCESS_CODE);
		data.setReturndata(true);
		//this.sendJson(resp, data);
	}
	
	
	
	
	protected void assemblePasteData(HttpServletRequest req,HttpServletResponse resp){
		JsonReturn data = new JsonReturn("");
		String step = req.getHeader("step");
		String excelId = req.getHeader("excelId");
		if(StringUtil.isEmpty(step)){
			step = "1";
		}
		Model model = new Model();
		int index = Integer.valueOf(step);
		model.setStep(index);
		model.setReqPath(29);
		model.setExcelId(excelId);
			//Thread.sleep(index*100);
		queueSender.send(Constant.queueName, model);
		data.setReturncode(Constant.SUCCESS_CODE);
		data.setReturndata(false);
		this.sendJson(resp, data);
	}
}
