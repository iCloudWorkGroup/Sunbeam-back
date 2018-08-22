package com.acmr.excel.controller;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLEncoder;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.springframework.context.annotation.Scope;
import org.springframework.data.mongodb.core.MongoTemplate;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.PathVariable;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.bind.annotation.RequestMethod;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.multipart.MultipartHttpServletRequest;
import org.springframework.web.servlet.ModelAndView;

import com.acmr.excel.controller.excelbase.BaseController;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.Position;
import com.acmr.excel.model.complete.CompleteExcel;
import com.acmr.excel.service.MBookService;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.FileUtil;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.UUIDUtil;
import com.acmr.excel.util.UploadThread;

import acmr.excel.ExcelException;
import acmr.excel.pojo.Constants.XLSTYPE;
import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColumn;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

/**
 * excel操作
 * 
 * @author jinhr
 */
 @Controller
 @RequestMapping
 @Scope("singleton")
public class ExcelController extends BaseController {
	private static Logger log = Logger.getLogger(ExcelController.class); 
	
	@Resource
	private BaseDao baseDao;
	@Resource
	private MSheetDao mexcelDao;
	
	@Resource
	private MBookService mbookService;
	
	@Resource
	private MongoTemplate mongoTemplate;
	
	

	/**
	 * excel下载
	 */
	@RequestMapping(value="/download/{excelId}",method=RequestMethod.GET)
	public void download(HttpServletRequest req, HttpServletResponse resp,@PathVariable String excelId) {
		String sheetId = excelId+0;
		ExcelBook excelBook = mbookService.getExcelBook(excelId, sheetId);
		if (excelBook != null) {
			try {
				OutputStream out = resp.getOutputStream();
				resp.setContentType("application/octet-stream");
				resp.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode("模板" + ".xlsx", "utf-8"));
				excelBook.saveExcel(out, XLSTYPE.XLSX);
				out.flush();
				out.close();
			} catch (UnsupportedEncodingException e) {
				e.printStackTrace();
			} catch (IOException e) {
				e.printStackTrace();
			} catch (ExcelException e) {
				e.printStackTrace();
			}
		}

	}
	
	@RequestMapping(value="/testup")
	public void testup(HttpServletRequest req, HttpServletResponse resp){
		
		List<MultipartFile> files = ((MultipartHttpServletRequest) req).getFiles("file");
		ExcelBook excel = new ExcelBook();
		InputStream is;
		try {
			is = files.get(0).getInputStream();
			if (ExcelUtil.isExcel2003(is)) {
				excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLS);
			} else {
				excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLSX);
			}
			ExcelSheet sheet = excel.getSheets().get(0);
			mongoTemplate.insert(sheet);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
		
	}
	
	@SuppressWarnings("unused")
	@RequestMapping(value="/testload",method=RequestMethod.GET)
	public void testload(HttpServletRequest req, HttpServletResponse resp){
		
		
			ExcelBook excel = null;
			if (excel != null) {
				try {
					OutputStream out = resp.getOutputStream();
					resp.setContentType("application/octet-stream");
					resp.setHeader("Content-disposition", "attachment;filename="+ URLEncoder.encode("模板" + ".xlsx", "utf-8"));
					excel.saveExcel(out, XLSTYPE.XLSX);
					out.flush();
					out.close();
				} catch (UnsupportedEncodingException e) {
					e.printStackTrace();
				} catch (IOException e) {
					e.printStackTrace();
				} catch (ExcelException e) {
					e.printStackTrace();
				}
			}
		
	}

	
	/**
	 * 初始化excel页面
	 */
	@RequestMapping(value="/main/{sbmId}")
	public ModelAndView main(HttpServletRequest req, HttpServletResponse resp,@PathVariable String sbmId) {
		
		log.info("初始化excel");
		
		String uuid = UUIDUtil.getUUID();
		
		return new ModelAndView("/index").addObject("sbmId", sbmId).
				addObject("frontName",Constant.frontName).addObject("uuid", uuid);
	}
	
	/**
	 * 初始化表明列表页面
	 */
	@RequestMapping("/")
	public ModelAndView welcom(HttpServletRequest req, HttpServletResponse resp) {
		List<String> excels = mbookService.getExcels();
		
		return new ModelAndView("/show").addObject("excels", excels);
	}
	

	/**
	 * 创建一个默认的excel
	 * 
	 * @param excelId
	 *            excelId
	 */
	private ExcelBook createTestExcel(String excelId) {
		ExcelBook excelBook = new ExcelBook();
		ExcelSheet sheet = new ExcelSheet();
		for (int i = 1; i < 27; i++) {
			ExcelColumn column = sheet.addColumn();
			column.setWidth(69);
		}
		for (int i = 1; i < 201; i++) {
			ExcelRow row = sheet.addRow();
			row.setHeight(19);
		}
		excelBook.getSheets().add(sheet);
		return excelBook;
	}
	private String readFile(String filepath) {
		String content = "";
		try {
			BufferedReader br = new BufferedReader(new FileReader(filepath));
			String str = null;
			StringBuffer buf = new StringBuffer();
			while ((str = br.readLine()) != null) {
				buf.append(str);
				buf.append("\r\n");
			}
			content = buf.toString();
			// System.out.println(content);
			br.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return content;
	}


	@RequestMapping(value="/reopen/{excelId}",method=RequestMethod.GET)
	public ModelAndView reopen(HttpServletRequest req, HttpServletResponse resp,@PathVariable String excelId) {

		return new ModelAndView("/index").addObject("excelId", excelId).addObject("sheetId", "1")
				.addObject("build", false).addObject("frontName",Constant.frontName);
	}


	
	/**
	 * 上传excel
	 * 
	 * @param req
	 * @param resp
	 * @throws Exception
	 */
	@RequestMapping("/upload")
	public void upload(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		List<MultipartFile> files = ((MultipartHttpServletRequest) req).getFiles("file");
		ExcelBook excel = new ExcelBook();
		InputStream is = files.get(0).getInputStream();
		if (ExcelUtil.isExcel2003(is)) {
			excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLS);
		} else {
			excel.LoadExcel(files.get(0).getInputStream(), XLSTYPE.XLSX);
		}
		
		/*ExcelSheet excelSheet = excel.getSheets().get(0);
		List<ExcelColumn> colList = excelSheet.getCols();
		int colSize = colList.size();
		if (colSize < 26) {
			for (int i = colSize; i < 26; i++) {
				excelSheet.addColumn();
			}
		}*/
		String excelId = UUIDUtil.getUUID();
		
		JsonReturn data = new JsonReturn("");
		
		boolean result =  mbookService.saveExcelBook(excel, excelId);
		
		if(result){
			data.setReturncode(200);
			data.setReturndata(excelId);
		}else{
			data = null;
			resp.setStatus(413);
		}
		
		this.sendJson(resp, data);
	}
	

	/**
	 * 重新打开excel
	 * 
	 * @throws Exception
	 */
	@RequestMapping(value="/reload")
	public void position(HttpServletRequest req, HttpServletResponse resp) throws Exception {
		String excelId = req.getHeader("X-Book-Id");
		String sheetId = excelId+0;
		Position position = getJsonDataParameter(req, Position.class);
		int height = position.getBottom();
		int right = position.getRight();
		
		CompleteExcel excel = mbookService.reload(excelId, sheetId, 0, height, 0, right,0);
	
		this.sendJson(resp, excel);
		
	}

	

	/**
	 * 大文件上传
	 * 
	 * @param req
	 * @param resp
	 */

	public void uploadBigFile(HttpServletRequest req, HttpServletResponse resp) {
		InputStream is;
		try {
			is = req.getInputStream();
			byte[] bytes = FileUtil.toByteArray(is);
			String name = req.getParameter("fname");
			ThreadPoolExecutor threadPool = (ThreadPoolExecutor) Executors
					.newCachedThreadPool();
			String partFileName = name + "." + System.currentTimeMillis()
					+ ".part";
			threadPool.execute(new UploadThread(partFileName, bytes));
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	/**
	 * 合并文件
	 * 
	 * @param req
	 * @param resp
	 */

	public void mergeFile(HttpServletRequest req, HttpServletResponse resp) {
		FileUtil fileUtil = new FileUtil();
		int blockFileSize = 1024 * 1024 * 10;
		String name = req.getParameter("fname");
		try {
			fileUtil.mergePartFiles(ExcelUtil.currentWorkDir, ".part",
					blockFileSize, ExcelUtil.currentWorkDir + name);
		} catch (IOException e) {
			e.printStackTrace();
		}
		JsonReturn data = new JsonReturn("");
		data.setReturncode(200);
		String address = "d:\\temp\\" + name;
		data.setReturndata(address);
		this.sendJson(resp, data);
	}

	/**
	 * 大文件下载
	 * 
	 * @param req
	 * @param resp
	 * @throws IOException
	 */

	public void downloadBigFile(HttpServletRequest req, HttpServletResponse resp)
			throws IOException {
		String fileName = req.getParameter("fname");
		InputStream is = new FileInputStream("d:\\temp\\" + fileName);
		resp.reset();
		resp.setContentType("application/pdf");
		resp.setHeader("Pragma", "public");
		resp.setHeader("Cache-Control", "max-age=30");
		resp.setHeader("Content-disposition", "inline;filename=" + fileName);
		ServletOutputStream out = resp.getOutputStream();
		ByteArrayOutputStream bos = new ByteArrayOutputStream();
		byte[] bytes = FileUtil.toByteArray(is);
		try {
			if (null != bytes) {
				bos.write(bytes);
				resp.setContentLength(bos.size());
				bos.writeTo(out);
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			out.close();
			out.flush();
			bos.close();
			bos.flush();
		}
	}

}
