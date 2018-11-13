package test;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.DVConstraint;
import org.apache.poi.hssf.usermodel.HSSFDataValidation;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hwpf.model.SavedByEntry;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.XSSFDataValidation;
import org.apache.poi.xssf.usermodel.XSSFDataValidationConstraint;
import org.apache.poi.xssf.usermodel.XSSFDataValidationHelper;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.jdbc.support.DatabaseStartupValidator;

import acmr.excel.pojo.ExcelBook;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelCellRangeAddress;
import acmr.excel.pojo.ExcelDataValidation;
import acmr.excel.pojo.ExcelRow;
import acmr.excel.pojo.ExcelSheet;

public class Test3 {
	public static void main(String[] args) throws Exception {

		
		 /*ExcelBook book = new ExcelBook(); book.LoadExcel("D://123.xlsx");
		 ExcelSheet sheet = book.getSheets().get(0); 
		  
		 book.saveExcel("D://456.xlsx"); 
		 ExcelRow row =sheet.getRows().get(0); 
		 List<ExcelCell> cells = row.getCells();*/
		 
		/*
		 * for (int i = 0; i < cells.size(); i++) { ExcelCell cell =
		 * cells.get(i); System.out.println(cell.getText() + ";" +
		 * cell.getValue() + ";" + cell.getType() + ";" +
		 * cell.getCellstyle().getDataformat()); }
		 */
		Test3 test = new Test3();
		test.saveExcel("xlsx");
	}

	public void saveExcel(String type) {

		FileOutputStream out = null;
		try {
			
			if("xlsx".equals(type)){
				// excel对象
				XSSFWorkbook wb = new XSSFWorkbook();
				// sheet对象
				XSSFSheet sheet = wb.createSheet("sheet1");
				setValidataXList(sheet);
				setValidataXInteger(sheet);
				setValidataXDecima(sheet);
				setValidataXText(sheet);
				// 输出excel对象
				out = new FileOutputStream("D://aaa1.xlsx");
				wb.write(out);
			}else{
				// excel对象
				HSSFWorkbook wb = new HSSFWorkbook();
				// sheet对象
				HSSFSheet sheet = wb.createSheet("sheet1");
				setValidata(sheet);
				setValidataList(sheet);
				setValidataInteger(sheet);
				setValidataDecima(sheet);
				setValidataText(sheet);
				setValidataDate(sheet);
				// 输出excel对象
				out = new FileOutputStream("D://aaa1.xls");
				wb.write(out);
			}
			
			out.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					// TODO Auto-generated catch block
					e.printStackTrace();
				}
			}
		}

	}


	public void setValidataX(XSSFSheet sheet) {
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
				.createExplicitListConstraint(
						new String[] { "11", "21", "31" });
		CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
		XSSFDataValidation validation = (XSSFDataValidation) dvHelper
				.createValidation(dvConstraint, addressList);
		validation.setSuppressDropDownArrow(true);
		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
	}
	
	public void setValidataXList(XSSFSheet sheet) {
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
				.createFormulaListConstraint("$D$1:$D$3");
		CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 2, 2);
		XSSFDataValidation validation = (XSSFDataValidation) dvHelper
				.createValidation(dvConstraint, addressList);
		validation.setSuppressDropDownArrow(true);
		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
	}
	
	public void setValidataXInteger(XSSFSheet sheet) {
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
				.createNumericConstraint(
					      XSSFDataValidationConstraint.ValidationType.INTEGER,
					      XSSFDataValidationConstraint.OperatorType.LESS_OR_EQUAL,
					      "10", "100");
		CellRangeAddressList addressList = new CellRangeAddressList(1, 1, 0, 0);
		XSSFDataValidation validation = (XSSFDataValidation) dvHelper
				.createValidation(dvConstraint, addressList);
		
		validation.setSuppressDropDownArrow(true);

		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
	}
	
	public void setValidataXDecima(XSSFSheet sheet) {
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
				.createNumericConstraint(
					      XSSFDataValidationConstraint.ValidationType.DECIMAL,
					      XSSFDataValidationConstraint.OperatorType.BETWEEN,
					      "10", "100");;
		CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 1, 1);
		XSSFDataValidation validation = (XSSFDataValidation) dvHelper
				.createValidation(dvConstraint, addressList);

		
		validation.setSuppressDropDownArrow(true);

		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
	}
	
	public void setValidataXText(XSSFSheet sheet) {
		XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
		XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint) dvHelper
				.createNumericConstraint(
					      XSSFDataValidationConstraint.ValidationType.TEXT_LENGTH,
					      XSSFDataValidationConstraint.OperatorType.BETWEEN,
					      "1", "5");;
		CellRangeAddressList addressList = new CellRangeAddressList(1, 1, 1, 1);
		XSSFDataValidation validation = (XSSFDataValidation) dvHelper
				.createValidation(dvConstraint, addressList);

		
		validation.setSuppressDropDownArrow(true);

		validation.setShowErrorBox(true);
		sheet.addValidationData(validation);
	}
	
	public void setValidata(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    0, 0, 0, 0);
			  DVConstraint dvConstraint = DVConstraint.createExplicitListConstraint(
			    new String[]{"10", "20", "30"});
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	
	public void setValidataList(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    0, 0, 2, 2);
			  DVConstraint dvConstraint = DVConstraint.createFormulaListConstraint("$C$1:$C$3");
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	
	public void setValidataInteger(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    1, 1, 0, 0);//行，行，列，列
			  DVConstraint dvConstraint = DVConstraint.createNumericConstraint(
					    DVConstraint.ValidationType.INTEGER,
					    DVConstraint.OperatorType.BETWEEN, "10", "100");
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	
	public void setValidataDecima(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    0, 0, 1, 1);//行，行，列，列
			  DVConstraint dvConstraint = DVConstraint.createNumericConstraint(
					    DVConstraint.ValidationType.DECIMAL,
					    DVConstraint.OperatorType.BETWEEN, "10", "100");
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	public void setValidataText(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    1, 1, 1, 1);//行，行，列，列
			  DVConstraint dvConstraint = DVConstraint.createNumericConstraint(
					    DVConstraint.ValidationType.TEXT_LENGTH,
					    DVConstraint.OperatorType.BETWEEN, "1", "9");
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	
	public void setValidataDate(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    0, 0, 2, 2);//行，行，列，列
			  DVConstraint dvConstraint = DVConstraint.createDateConstraint(
					    DVConstraint.ValidationType.DATE,"2010/1/2", "2010/1/6","yyyy/M/d");
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}
	
	public void setValidataTime(HSSFSheet sheet) {
		CellRangeAddressList addressList = new CellRangeAddressList(
			    0, 0, 2, 2);//行，行，列，列
			  DVConstraint dvConstraint = DVConstraint.createTimeConstraint( DVConstraint.ValidationType.TIME, "", "");
					   
			  DataValidation dataValidation = new HSSFDataValidation
			    (addressList, dvConstraint);
			  dataValidation.setSuppressDropDownArrow(false);
			  sheet.addValidationData(dataValidation);
	}

}
