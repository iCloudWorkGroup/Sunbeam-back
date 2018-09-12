package com.acmr.excel.util;

import java.math.BigDecimal;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFDataFormat;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.Format;
import com.acmr.excel.model.mongo.MCell;

import acmr.excel.pojo.Constants.CELLTYPE;
import acmr.excel.pojo.ExcelCell;
import acmr.excel.pojo.ExcelColor;
import acmr.excel.pojo.ExcelFormat;

public class CellFormateUtil {
	/**
	 * 文本类型
	 * 
	 * @param excelCell
	 */
	public static void setText(ExcelCell excelCell) {
		// excelCell.setShowText(excelCell.getText());
		excelCell.getExps().put("displayText", excelCell.getText());
		excelCell.setType(CELLTYPE.STRING);
		excelCell.getCellstyle().setDataformat("@");
		excelCell.setValue(excelCell.getText());
	}

	/**
	 * 常规类型
	 * 
	 * @param excelCell
	 */
	public static void setGeneral(ExcelCell excelCell) {
		String text = excelCell.getText();
		excelCell.getExps().put("displayText", text);
		// excelCell.setShowText(text);
		if (isNumeric(text)) {
			excelCell.setType(CELLTYPE.NUMERIC);
		} else {
			excelCell.setType(CELLTYPE.STRING);
		}
		excelCell.getCellstyle().setDataformat("General");
	}

	/**
	 * 数字类型
	 * 
	 * @param excelCell
	 * @param decimalPoint
	 * @param thousandPoint
	 */
	// public static void setNumber(ExcelCell excelCell,int decimalPoint,boolean
	// thousandPoint){
	// String text = excelCell.getText();
	//// if (isNumeric(text)) {
	//// excelCell.setValue(Double.valueOf(text));
	//// //小数位
	//// text = setDecimalPoint(decimalPoint, text);
	//// //千分位
	//// text = setThousandPoint(thousandPoint, text);
	//// //excelCell.setShowText(text);
	//// }
	// if(!StringUtil.isEmpty(text)){
	// excelCell.setType(CELLTYPE.NUMERIC);
	// }
	// excelCell.getCellstyle().setDataformat(getNumDataFormate(decimalPoint,
	// thousandPoint));
	// excelCell.getExps().put("decimalPoint", decimalPoint+"");
	// excelCell.getExps().put("thousandPoint", thousandPoint+"");
	// }

	public static String setNumber(ExcelCell excelCell, int decimalPoint,
			boolean thousandPoint) {
		String text = excelCell.getText();
		String result = "";
		if (isNumeric(text)) {
			excelCell.setValue(Double.valueOf(text));
			// 小数位
			text = setDecimalPoint(decimalPoint, text);
			// 千分位
			text = setThousandPoint(thousandPoint, text);
			result = text;
			// excelCell.setShowText(text);
			excelCell.getExps().put("displayText", text);
			excelCell.setType(CELLTYPE.NUMERIC);
		}
		excelCell.getCellstyle()
				.setDataformat(getNumDataFormate(decimalPoint, thousandPoint));
		excelCell.getExps().put("decimalPoint", decimalPoint + "");
		excelCell.getExps().put("thousandPoint", thousandPoint + "");

		return result;
	}

	/**
	 * 日期类型
	 * 
	 * @param excelCell
	 */
	public static void setTime(ExcelCell excelCell, String dataFormate) {
		String text = excelCell.getText();
		String textFormate = "";
		if (text.contains("年") && text.contains("月") && text.contains("日")) {
			textFormate = "yyyy年MM月dd日";
		} else if (text.contains("年") && text.contains("月")
				&& !text.contains("日")) {
			textFormate = "yyyy年MM月";
		} else if (text.contains("/")) {
			textFormate = "yyyy/MM/dd";
		}
		String outputFormate = null;
		switch (dataFormate) {
		case "yyyy年MM月dd日":
			outputFormate = "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy";
			break;
		case "yyyy年MM月":
			outputFormate = "yyyy\"年\"m\"月\";@";
			break;
		case "yyyy/MM/dd":
			outputFormate = "m/d/yy";
			break;

		default:
			break;
		}
		SimpleDateFormat format = null;
		switch (textFormate) {
		case "yyyy年MM月dd日":
			format = new SimpleDateFormat("yyyy年MM月dd日");
			try {
				format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				// excelCell.setShowText(newDate);
				excelCell.getExps().put("displayText", newDate);
				excelCell.setType(CELLTYPE.DATE);
				if (outputFormate != null) {
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				// e.printStackTrace();
			}
			break;
		case "yyyy年MM月":
			format = new SimpleDateFormat("yyyy年MM月");
			try {
				format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				// excelCell.setShowText(newDate);
				excelCell.getExps().put("displayText", newDate);
				excelCell.setType(CELLTYPE.DATE);
				if (outputFormate != null) {
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				// e.printStackTrace();
			}
			break;
		case "yyyy/MM/dd":
			format = new SimpleDateFormat("yyyy/MM/dd");
			try {
				format.setLenient(false);
				Date date = format.parse(text);
				format = new SimpleDateFormat(dataFormate);
				String newDate = format.format(date);
				excelCell.setValue(date);
				// excelCell.setShowText(newDate);
				excelCell.getExps().put("displayText", newDate);
				excelCell.setType(CELLTYPE.DATE);
				if (outputFormate != null) {
					excelCell.getCellstyle().setDataformat(outputFormate);
				}
				excelCell.getExps().put("dataFormate", dataFormate);
			} catch (ParseException e) {
				// e.printStackTrace();
			}
			break;

		default:
			break;
		}
	}

	/**
	 * 货币类型
	 * 
	 * @param excelCell
	 * @param decimalPoint
	 * @param currencySymbol
	 */
	public static String setCurrency(ExcelCell excelCell, int decimalPoint,
			String currencySymbol) {
		String text = excelCell.getText();
		String result = "";
		if (isNumeric(text)) {
			excelCell.setValue(Double.valueOf(text));
			// 小数位
			text = setDecimalPoint(decimalPoint, text);
			// 千分位
			text = setThousandPoint(true, text);
			result = currencySymbol + text;
			// excelCell.setShowText(currencySymbol + text);
			// excelCell.getExps().put("displayText", currencySymbol + text);
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat(
					getCurrencyDataFormate(decimalPoint, currencySymbol));
			excelCell.getExps().put("decimalPoint", decimalPoint + "");
			excelCell.getExps().put("currencySymbol", currencySymbol);
		}
		return result;
	}

	/**
	 * 百分比
	 * 
	 * @param excelCell
	 * @param decimalPoint
	 */
	public static void setPercent(ExcelCell excelCell, int decimalPoint) {
		String text = excelCell.getText();
		if (isNumeric(text)) {
			double num = Double.valueOf(text);
			excelCell.setValue(num);
			num *= 100;
			text = setDecimalPoint(decimalPoint, num + "");
			// excelCell.setShowText(text + "%");
			excelCell.getExps().put("displayText", text + "%");
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle()
					.setDataformat(getPercentDataFormate(decimalPoint));
		}
	}

	/**
	 * 日期校验
	 * 
	 * @param date
	 * @param pattern
	 * @return
	 */
	public static Date getDate(String date, String pattern) {
		// boolean convertSuccess = true;
		Date retDate = null;
		SimpleDateFormat format = new SimpleDateFormat(pattern);
		try {
			format.setLenient(false);
			retDate = format.parse(date);
		} catch (ParseException e) {
			// convertSuccess = false;
		}
		return retDate;
	}

	/**
	 * 设置数字格式
	 * 
	 * @param decimalPoint
	 * @param thousandPoint
	 * @return
	 */
	public static String getNumDataFormate(int decimalPoint,
			boolean thousandPoint) {
		String dataFormate = "";
		if (thousandPoint) {
			dataFormate += "#,##";
		}
		dataFormate += "0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "_);\\(" + dataFormate + "\\)";
		return dataFormate;
	}

	/**
	 * 设置货币格式
	 * 
	 * @param decimalPoint
	 * @return
	 */
	public static String getCurrencyDataFormate(int decimalPoint,
			String currencySymbol) {
		String dataFormate = "\"" + currencySymbol + "\"#,##0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "_);\\(" + dataFormate + "\\)";
		return dataFormate;
	}

	/**
	 * 设置百分比格式
	 * 
	 * @param decimalPoint
	 * @return
	 */
	public static String getPercentDataFormate(int decimalPoint) {
		String dataFormate = "0";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				dataFormate += ".";
			}
			dataFormate += "0";
		}
		dataFormate += "%";
		return dataFormate;
	}

	/**
	 * 设置小数位
	 * 
	 * @param decimalPoint
	 * @param text
	 * @return
	 */
	public static String setDecimalPoint(int decimalPoint, String text) {
		if (decimalPoint < 0) {
			return null;
		}
		String temp = "#";
		for (int i = 0; i < decimalPoint; i++) {
			if (i == 0) {
				temp += ".";
			}
			temp += "0";
		}
		DecimalFormat df = new DecimalFormat(temp);
		return df.format(Double.valueOf(text));
	}

	/**
	 * 设置千分位
	 * 
	 * @param thousandPoint
	 * @param text
	 * @return
	 */
	public static String setThousandPoint(boolean thousandPoint, String text) {
		String retVal = null;
		NumberFormat numberFormat = NumberFormat.getNumberInstance();
		if (thousandPoint) {
			// 千分位
			numberFormat.setGroupingUsed(true);
		} else {
			numberFormat.setGroupingUsed(false);
		}
		if (text.indexOf(".") < 0) {
			// 整数
			retVal = numberFormat.format(Long.valueOf(text));
		} else {
			int textLength = text.split("\\.")[1].length();
			// 小数
			numberFormat.setMinimumFractionDigits(textLength);
			retVal = numberFormat.format(Double.valueOf(text));
			String[] news = retVal.split("\\.");
			int retLength = 0;
			if (news != null && news.length > 1) {
				retLength = news[1].length();
			} else {
				retVal += ".";
			}
			if (textLength > retLength) {
				for (int i = retLength; i < textLength; i++) {
					retVal += 0;
				}
			}
		}
		return retVal;
	}

	/**
	 * 判断是否为数字
	 * 
	 * @param str
	 * @return
	 */
	public static boolean isNumeric(String str) {
		// [-+]?\\d*\\.?\\d+, ^[0-9]+(\\.[0-9]+){0,1}$
		// Pattern pattern = Pattern.compile("-?[0-9]+[.]?[0-9]+");
		Pattern pattern = Pattern.compile("[-+]?\\d*\\.?\\d+");
		Matcher isNum = pattern.matcher(str);
		if (!isNum.matches()) {
			return false;
		}
		return true;
		// return NumberUtils.isNumber(str);
	}

	/**
	 * 自动识别
	 * 
	 * @param content
	 */

	public static void autoRecognise(String content, ExcelCell excelCell) {
		String pattern1 = "yyyy年MM月dd日";
		String pattern2 = "yyyy年MM月";
		String pattern3 = "yyyy/MM/dd";
		Date d1 = getDate(content, pattern1);
		Date d2 = getDate(content, pattern2);
		Date d3 = getDate(content, pattern3);
		if (isNumeric(content)) {
			content = getPrettyNumber(content);
			excelCell.setType(CELLTYPE.NUMERIC);
			excelCell.getCellstyle().setDataformat("General");
			excelCell.setValue(Double.valueOf(content));
			// excelCell.setShowText(content);
			excelCell.getExps().put("displayText", content);
		} else if (d1 != null) {
			excelCell.setValue(d1);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle()
					.setDataformat("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");
			// excelCell.setShowText(content);
			excelCell.getExps().put("displayText", content);
		} else if (d2 != null) {
			excelCell.setValue(d2);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle().setDataformat("yyyy\"年\"m\"月\";@");
			// excelCell.setShowText(content);
			excelCell.getExps().put("displayText", content);
		} else if (d3 != null) {
			excelCell.setValue(d3);
			excelCell.setType(CELLTYPE.DATE);
			excelCell.getCellstyle().setDataformat("m/d/yy");
			// excelCell.setShowText(content);
			excelCell.getExps().put("displayText", content);
		}
	}

	/**
	 * 自动识别
	 * 
	 * @param content
	 */

	public static void autoRecognise(String text, MCell mcell) {
		String pattern1 = "yyyy年MM月dd日";
		String pattern2 = "yyyy年MM月";
		String pattern3 = "yyyy/MM/dd";
		Date d1 = getDate(text, pattern1);
		Date d2 = getDate(text, pattern2);
		Date d3 = getDate(text, pattern3);
		Content content = mcell.getContent();
		if (isNumeric(text)) {
			text = getPrettyNumber(text);
			content.setType(CELLTYPE.NUMERIC.name());
			content.setExpress("General");

		} else if (d1 != null) {
			content.setType(CELLTYPE.DATE.name());
			content.setExpress("[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy");

		} else if (d2 != null) {
			content.setType(CELLTYPE.DATE.name());
			content.setExpress("yyyy\"年\"m\"月\";@");

		} else if (d3 != null) {
			content.setType(CELLTYPE.DATE.name());
			content.setExpress("m/d/yy");

		}
	}

	public static void setShowText(Content content) {
		try {
			String type = content.getType();
			if ("date".equals(type)) {
				String text = content.getTexts();
				String express = content.getExpress();
				String pattern2 = "yyyy/MM/dd";
				String pattern1 = "yyyy年MM月dd日";
				String pattern3 = "yyyy/M/d";
				String pattern4 = "yyyy年M月d日";
				if ("yyyy/mm/dd".equals(express)) {
					Date d1 = getDate(text, pattern1);
					if (null == d1) {
						Date d3 = getDate(text,pattern3);
						if(null==d3){
							content.setDisplayTexts(text);
							return;
						}else{
							SimpleDateFormat format = new SimpleDateFormat(pattern2);
							String display = format.format(d3);
							content.setDisplayTexts(display);
							return;
						}
					} else {
						SimpleDateFormat format = new SimpleDateFormat(pattern2);
						String display = format.format(d1);
						content.setDisplayTexts(display);
						return;
					}
				}else{
					Date d1 = getDate(text, pattern2);
					if (null == d1) {
						Date d4 = getDate(text,pattern4);
						if(null==d4){
							content.setDisplayTexts(text);
							return;
						}else{
							SimpleDateFormat format = new SimpleDateFormat(pattern1);
							String display = format.format(d4);
							content.setDisplayTexts(display);
							return;
						}
					} else {
						SimpleDateFormat format = new SimpleDateFormat(pattern1);
						String display = format.format(d1);
						content.setDisplayTexts(display);
						return;
					}
				}
			}
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet = wb.createSheet();
			XSSFRow row = sheet.createRow(0);
			XSSFCell cell = row.createCell(0);

			switch (type) {
			case "routine":
			case "text":
				cell.setCellValue(content.getTexts());
				break;
			case "number":
			case "currency":
			case "percent":
				try {
					cell.setCellValue(Double.valueOf(content.getTexts()));
				} catch (Exception e) {
					cell.setCellValue(content.getTexts());
				}
				break;
			default:
				break;
			}
			XSSFCellStyle style = cell.getCellStyle();
			XSSFDataFormat format = wb.createDataFormat();
			style.setDataFormat(format.getFormat(content.getExpress()));
			cell.setCellStyle(style);
			DataFormatter f = new DataFormatter();
			String text = f.formatCellValue(cell);
			if ("".equals(text) || null == text) {
				content.setDisplayTexts(content.getTexts());
			} else {
				content.setDisplayTexts(text);
			}
			
		} catch (Exception e) {
			content.setDisplayTexts(content.getTexts());
		}
	
	}

	public static void setShowText(ExcelCell excelCell, Content content) {
		if (CELLTYPE.BLANK == excelCell.getType()) {
			content.setDisplayTexts("");
			return;
		}
		String dataFormate = excelCell.getCellstyle().getDataformat();
		String text = excelCell.getText();
		Object o = excelCell.getValue();
		SimpleDateFormat sdf = null;
		switch (dataFormate) {
		case "General":
			// NUMERIC
			if (CELLTYPE.NUMERIC == excelCell.getType()) {

				if (o != null) {
					DecimalFormat df = new DecimalFormat("#.######");
					String value = df.format(o);
					if (!text.equals(o.toString())) {
						content.setTexts(value);
					}
					text = value;
				}
			}

			content.setDisplayTexts(text);
			break;
		case "@":
			content.setTexts(text);
			content.setDisplayTexts(text);
			break;
		case "[$-F800]dddd\\,\\ mmmm\\ dd\\,\\ yyyy":
			sdf = new SimpleDateFormat("yyyy年MM月dd日");
			try {
				content.setDisplayTexts(sdf.format(excelCell.getValue()));
			} catch (Exception e) {
				content.setDisplayTexts((String) o);
			}

			content.setTexts(excelCell.getText());

			break;
		case "yyyy\"年\"m\"月\";@":
			sdf = new SimpleDateFormat("yyyy年MM月");
			try {
				content.setDisplayTexts(sdf.format(o));
			} catch (Exception e) {
				content.setDisplayTexts((String) o);
			}

			content.setTexts(content.getDisplayTexts());

			break;
		case "m/d/yy":
			sdf = new SimpleDateFormat("yyyy/MM/dd");
			try {
				content.setDisplayTexts(sdf.format(o));
			} catch (Exception e) {
				content.setDisplayTexts((String) o);
			}

			content.setTexts(content.getDisplayTexts());

			break;
		case "0_);\\(0\\)":
		case "0_ ":
		case "0_);[Red]\\(0\\)":
		case "0;[Red]0":
		case "0_ ;[Red]\\-0\\ ":
			// 整数

			content.setTexts(o.toString());
			String zdisplayText = setNumber(excelCell, 0, false);
			content.setDisplayTexts(zdisplayText);

			break;
		case "#,##0_);\\(#,##0\\)":
			// 整数千分位
			setNumber(excelCell, 0, true);
			content.setDisplayTexts(excelCell.getShowText());

			break;
		case "#,##0_);[Red]\\(#,##0\\)":
			// 整数千分位
			setNumber(excelCell, 0, true);
			content.setDisplayTexts(excelCell.getShowText());

			break;
		case "\"¥\"#,##0_);\\(\"¥\"#,##0\\)":
			// 货币
			setCurrency(excelCell, 0, "¥");

			content.setDisplayTexts(excelCell.getShowText());

			break;
		case "\"$\"#,##0_);\\(\"¥\"#,##0\\)":
			// 货币
			setCurrency(excelCell, 0, "$");

			content.setDisplayTexts(excelCell.getShowText());

			break;
		case "0%":
			setPercent(excelCell, 0);

			content.setDisplayTexts(excelCell.getShowText());

			break;
		default:
			if (dataFormate.startsWith("0.0") && dataFormate.endsWith("0\\)")
					&& !dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				String d = setNumber(excelCell, index - 2, false);

				content.setDisplayTexts(d);
				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
			} else if (dataFormate.startsWith("0.0")) {
				// 小数
				String d = setNumber(excelCell, 1, false);

				content.setDisplayTexts(d);
				content.setTexts(d);
				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
			} else if (dataFormate.startsWith("0.0")
					&& dataFormate.endsWith("0\\)")
					&& dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				String displayText = setNumber(excelCell, index - 2, false);

				content.setDisplayTexts(displayText);
				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
				content.setDisplayTexts(displayText);
				// String showText = excelCell.getShowText();
				// if(showText != null && showText.contains("-")){
				// showText = showText.replace("-", "");
				// content.setDisplayTexts("("+showText+")");
				// //ExcelColor excelColor =
				// excelCell.getCellstyle().getFont().getColor();
				// //excelColor.setAuto(false);
				// //excelColor.setRGB(255, 0, 0);
				// }else{
				// content.setDisplayTexts(showText);
				// }

			} else if (dataFormate.startsWith("0.0")
					&& !dataFormate.contains(";[Red]")
					&& dataFormate.indexOf("_") != -1) {
				// 小数
				int index = dataFormate.indexOf("_");
				int decimal = index - 2;
				String displayText = setNumber(excelCell, index - 2, false);

				content.setDisplayTexts(displayText);
				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
			} else if (dataFormate.startsWith("0.0")
					&& dataFormate.contains(";[Red]")) {
				// 小数
				int index = dataFormate.indexOf(";");
				int decimal = index - 2;
				String displayText = setNumber(excelCell, index - 2, false);

				content.setDisplayTexts(displayText);
				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
			} else if (dataFormate.startsWith("#,##0.0")) {
				// 小数千分位
				int index = dataFormate.indexOf("_");
				int decimal = index - 6;
				String displayText = setNumber(excelCell, decimal, true);
				content.setDisplayTexts(displayText);

				if (CELLTYPE.STRING == excelCell.getType()) {

					content.setDisplayTexts(excelCell.getText());
				}
			} else if (dataFormate.startsWith("\"¥\"#,##0.0")) {
				// 千分位货币
				int index = dataFormate.indexOf("_");
				int decimal = index - 9;
				String cu = setCurrency(excelCell, decimal, "¥");
				content.setDisplayTexts(cu);

			} else if (dataFormate.startsWith("\"$\"#,##0.0")) {
				// 千分位货币
				int index = dataFormate.indexOf("_");
				int decimal = index - 9;
				String cu2 = setCurrency(excelCell, index - 9, "$");
				content.setDisplayTexts(cu2);

			} else if (dataFormate.startsWith("0.0")
					&& dataFormate.endsWith("0%")) {
				// 小数百分比
				int index = dataFormate.indexOf("%");
				int decimal = index - 2;
				setPercent(excelCell, index - 2);

				content.setDisplayTexts(excelCell.getShowText());

			} else {

				if (CELLTYPE.NUMERIC == excelCell.getType()) {
					int d = ExcelFormat.getDecimalFormatDotcount(
							excelCell.getCellstyle().getDataformat());

					if (o != null) {
						text = o.toString();
						content.setTexts(text);
					}
				}
				content.setDisplayTexts(excelCell.getText());

			}
			break;
		}
	}

	private static String getPrettyNumber(String number) {
		return BigDecimal.valueOf(Double.parseDouble(number))
				.stripTrailingZeros().toPlainString();
	}

	public static String getBackground(ExcelColor color) {
		String s = null;
		if (null != color) {
			s = "rgb(" + color.getR() + "," + color.getG() + "," + color.getB()
					+ ")";
		} else {
			s = "rgb(0,0,0)";
		}
		return s;
	}

	public static String TypeToMtype(CELLTYPE type) {
		String mtype = "";
		switch (type) {
		case BLANK:
			mtype = "routine";
			break;
		case NUMERIC:
			mtype = "number";
			break;
		case STRING:
			mtype = "text";
			break;
		case DATE:
			mtype = "date";
			break;
		default:
			mtype = "routine";
			break;
		}
		return mtype;
	}

	public static CELLTYPE MTypeToType(String mtype) {
		CELLTYPE type;
		switch (mtype) {
		case "routine":
			type = CELLTYPE.STRING;
			break;
		case "text":
			type = CELLTYPE.STRING;
			break;
		case "number":
			type = CELLTYPE.NUMERIC;
			break;
		case "date":
			type = CELLTYPE.DATE;
			break;
		case "currency":
			type = CELLTYPE.NUMERIC;
			break;
		case "percent":
			type = CELLTYPE.NUMERIC;
			break;
		default:
			type = CELLTYPE.STRING;
			break;
		}
		return type;
	}

	public static void main(String[] args) throws ParseException {
		/*
		 * XSSFWorkbook wb = new XSSFWorkbook(); XSSFSheet sheet =
		 * wb.createSheet(); XSSFRow row = sheet.createRow(0); XSSFCell cell =
		 * row.createCell(0);
		 * 
		 * SimpleDateFormat simpleDateFormat=new SimpleDateFormat("yyyy/MM/dd");
		 * Date date = simpleDateFormat.parse("2018/8/8");
		 * 
		 * cell.setCellValue("2018/8/8"); XSSFCellStyle style =
		 * cell.getCellStyle(); XSSFDataFormat format= wb.createDataFormat();
		 * style.setDataFormat(format.getFormat("yyyy\"年\"m\"月\"d\"日\";@"));
		 * cell.setCellStyle(style); DataFormatter f = new DataFormatter();
		 * String text = f.formatCellValue(cell); System.out.println(text);
		 */

		String pattern1 = "yyyy年MM月dd日";
		String pattern2 = "yyyy年MM月";
		String pattern3 = "yyyy/MM/dd";
		String content = "2010/3/8";
		Date d1 = getDate(content, pattern1);
		Date d2 = getDate(content, pattern2);
		Date d3 = getDate(content, pattern3);
		System.out.println(d3);

	}
}
