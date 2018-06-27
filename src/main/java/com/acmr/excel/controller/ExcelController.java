package com.acmr.excel.controller;

import java.io.BufferedReader;
import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.concurrent.Executors;
import java.util.concurrent.ThreadPoolExecutor;

import javax.annotation.Resource;
import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;

import org.apache.log4j.Logger;
import org.springframework.context.annotation.Scope;
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
import com.acmr.excel.model.history.VersionHistory;
import com.acmr.excel.service.MBookService;
import com.acmr.excel.service.MSheetService;
import com.acmr.excel.util.ExcelUtil;
import com.acmr.excel.util.FileUtil;
import com.acmr.excel.util.JsonReturn;
import com.acmr.excel.util.UUIDUtil;
import com.acmr.excel.util.UploadThread;

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

	/**
	 * excel下载
	 */
	@RequestMapping(value="/download/{excelId}",method=RequestMethod.GET)
	public void download(@PathVariable String excelId,HttpServletRequest req, HttpServletResponse resp) {
		/*ExcelBook excelBook;
		if (excelBook != null) {
			//excelService.changeHeightOrWidth(excelBook);
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
		}*/

	}

	
	/**
	 * 初始化excel页面
	 */
	@RequestMapping(value="/main/{sbmId}")
	public ModelAndView main(@PathVariable String sbmId) {
		//handleExcelService.createNewExcel(excelId,mongodbServiceImpl);
		//VersionHistory versionHistory = new VersionHistory();
		//storeService.set(excelId+"_history", versionHistory);
		log.info("初始化excel");
		// ExcelBook e = (ExcelBook)memcachedClient.get(excelId);
		// } <input type="hidden" id="excelId" value="(.*)"/>
		return new ModelAndView("/index").addObject("sbmId", sbmId).
				addObject("frontName",Constant.frontName);
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
	 * 测试接口
	 * 
	 * @param req
	 * @param resp
	 * @throws InterruptedException 
	 */

	public void test(HttpServletRequest req, HttpServletResponse resp) throws InterruptedException {
		String flag = req.getParameter("flag");
		String excelId = UUIDUtil.getUUID();
		ExcelBook excelBook = createTestExcel(excelId);
		ExcelSheet excelsheet = excelBook.getSheets().get(0);
		List<ExcelRow> rowList = excelsheet.getRows();
		for (int i = 0; i < 200; i++) {
			List<ExcelCell> excelList = rowList.get(i).getCells();
			for (int j = 0; j < 25; j++) {
				ExcelCell excelCell = excelList.get(j);
				if (excelCell == null) {
					excelCell = new ExcelCell();
				}
				excelCell.setText("回fff的痕迹卡的很金卡号地块和电视剧阿卡");
				excelCell.setValue("回到fffff的痕迹卡的很金卡号地块和电视剧阿卡");
				excelList.set(j, excelCell);
			}
		}
		//excelsheet.MergedRegions(4, 4, 6, 6);
		//storeService.set(excelId, excelBook);
		//storeService.set(excelId+"_ope", 0);
		JsonReturn data = new JsonReturn("");
		data.setReturndata(excelId);
		this.sendJson(resp, data);
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
	
	/**
	 * 前台获得js文件
	 * 
	 * @param req
	 * @param resp
	 */
	@RequestMapping(value="/getscript/{excelId}",method=RequestMethod.GET)
	public void getscript(@PathVariable String excelId,HttpServletRequest req, HttpServletResponse resp) {
		

	}
	@RequestMapping(value="/getscript",method=RequestMethod.GET)
	public void getscripts(HttpServletRequest req, HttpServletResponse resp) {
		//String excelId = req.getParameter("excelId");
		getscript(null, req, resp);

	}

	@RequestMapping(value="/reopen/{excelId}",method=RequestMethod.GET)
	public ModelAndView reopen(@PathVariable String excelId) {
		//String excelId = req.getParameter("excelId");
		VersionHistory versionHistory = new VersionHistory();
		//storeService.set(excelId+"_history",  versionHistory);
		return new ModelAndView("/index").addObject("excelId", excelId).addObject("sheetId", "1")
				.addObject("build", false).addObject("frontName",Constant.frontName);
	}

	/**
	 * 通过别名加载excel
	 */

	public void openExcelByAlais(HttpServletRequest req,HttpServletResponse resp) throws IOException {
		/*String excelId = req.getParameter("excelId");
		// int sheetId = Integer.valueOf(req.getParameter("sheetId"));
		String rowBegin = req.getParameter("rowBeginAlais");
		String rowEnd = req.getParameter("rowEndAlais");
		ExcelBook excelBook = (ExcelBook)  mongodbServiceImpl.get(null, null, null);
		// Workbook workbook = this.mockWorkbook();
		// CompleteExcel excel = excelService.getExcel(workbook);
		JsonReturn data = new JsonReturn("");
		if (excelBook != null) {
			ExcelSheet excelSheet = excelBook.getSheets().get(0);
			ReturnParam returnParam = new ReturnParam();
			CompleteExcel excel = new CompleteExcel();
			SpreadSheet spreadSheet = new SpreadSheet();
			excel.getSheets().add(spreadSheet);
			spreadSheet = excelService.openExcelByAlais(spreadSheet,
					excelSheet, rowBegin, rowEnd, returnParam);
			data.setReturncode(Constant.SUCCESS_CODE);
			data.setReturndata(excel);
			data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
		} else {
			data.setReturncode(Constant.CACHE_INVALID_CODE);
			data.setReturndata(Constant.CACHE_INVALID_MSG);
		}
		// System.out.println("openexcel====================="+JSON.toJSONString(data));
		this.sendJson(resp, data);*/
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
	 * 上传完成之后的页面
	 * @param req
	 * @param resp
	 * @return
	 * @throws Exception
	 */
//	@RequestMapping("/uploadComplete/{excelId}")
//	public ModelAndView uploadComplete(@PathVariable String excelId) throws Exception {
//		return new ModelAndView("/index").addObject("excelId", excelId).addObject("frontName",Constant.frontName);
//	}


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

	// /**
	// * 带冻结的定位还原
	// * @param req
	// * @param resp
	// * @throws Exception
	// */
	// public void positionWithFrozen(HttpServletRequest req,
	// HttpServletResponse resp) throws Exception{
	// String excelId = req.getParameter("excelId");
	// String height = req.getParameter("height");
	// String width = req.getParameter("width");
	// height = "800";
	// width = "300";
	// CompleteExcel excel =
	// (CompleteExcel)getSession(req).getAttribute(excelId);
	// Workbook workbook = this.mockWorkbook();
	// excel = excelService.getExcel(workbook);
	// SheetElement sheet = excel.getSpreadSheet().get(0).getSheet();
	// StartAlais startAlaisTest = sheet.getStartAlais();
	// sheet.getFrozen().setColIndex("3");
	// sheet.getFrozen().setRowIndex("22");
	// sheet.getFrozen().setDisplayAreaStartAlaisX("3");
	// sheet.getFrozen().setDisplayAreaStartAlaisY("22");
	// sheet.getFrozen().setState("1");
	// int glyLength = sheet.getGlY().size();
	// Gly gly = sheet.getGlY().get(glyLength-1);
	// int maxPixel = gly.getHeight() + gly.getTop();
	// startAlaisTest.setAlaisX("1");
	// startAlaisTest.setAlaisY("19");
	// ReturnParam returnParam = new ReturnParam();
	// if(excel != null){
	// excel = excelService.positionExcelWithFrozen(excel,
	// height,width,returnParam);
	// }
	// JsonReturn data = new JsonReturn("");
	// data.setReturncode(200);
	// data.setReturndata(excel);
	// // data.setxStartAlaisIndex(returnParam.getxStartAlaisIndex());
	// // data.setxEndAlaisIndex(returnParam.getxEndAlaisIndex());
	// // data.setyStartAlaisIndex(returnParam.getyStartAlaisIndex());
	// // data.setyEndAlaisIndex(returnParam.getyEndAlaisIndex());
	// data.setDataColStartIndex(returnParam.getDataColStartIndex());
	// data.setDataRowStartIndex(returnParam.getDataRowStartIndex());
	// data.setDisplayColStartAlias(returnParam.getDisplayColStartAlias());
	// data.setDisplayRowStartAlias(returnParam.getDisplayRowStartAlias());
	// data.setMaxPixel(maxPixel);
	// String jsonData = JSON.toJSONString(data);
	// ////system.out.println(jsonData);
	// this.sendJson(resp, data);
	// }
	/**
	 * 保存excel(关闭浏览器时的操作)
	 * 
	 * @throws Exception
	 */

	public void close(HttpServletRequest req, HttpServletResponse resp){
	}
	// public void test(HttpServletRequest request,HttpServletResponse
	// response){
	// this.memcachedClient.add("a", 7200, "aaaaaaaaaaaa");
	// String result = this.memcachedClient.get("a").toString();
	// System.out.println(result);
	// }

	// public void uploadBigFile(HttpServletRequest req, HttpServletResponse
	// resp){
	// List<MultipartFile> files = ((MultipartHttpServletRequest)
	// req).getFiles("file");
	// if (!files.isEmpty() && files.size() > 0) {
	// ThreadPoolExecutor threadPool = (ThreadPoolExecutor)
	// Executors.newCachedThreadPool();
	// for (int i = 0; i < files.size(); i++) {
	// MultipartFile file = files.get(i);
	// String partFileName = file.getName() + "." + (i+1) + ".part";
	// try {
	// threadPool.execute(new UploadThread(partFileName, file.getBytes()));
	// } catch (IOException e) {
	// e.printStackTrace();
	// }
	// }
	// }
	// }

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
