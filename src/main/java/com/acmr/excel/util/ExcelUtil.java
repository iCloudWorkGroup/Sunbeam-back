package com.acmr.excel.util;



import org.apache.poi.POIXMLDocument;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import acmr.excel.pojo.ExcelColor;


import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.ObjectInputStream;
import java.io.ObjectOutputStream;
import java.io.OutputStream;
import java.io.PushbackInputStream;


public class ExcelUtil {

	public static String currentWorkDir = "d:\\temp\\";

	/**
	 * 用序列化与反序列化实现深克隆
	 * 
	 * @param src
	 *            Object对象
	 * @return 克隆对象
	 */
	public static Object deepClone(Object src) {
		Object object = null;
		try {
			if (src != null) {
				ByteArrayOutputStream baos = new ByteArrayOutputStream();
				ObjectOutputStream oos = new ObjectOutputStream(baos);
				oos.writeObject(src);
				oos.close();
				ByteArrayInputStream bais = new ByteArrayInputStream(
						baos.toByteArray());
				ObjectInputStream ois = new ObjectInputStream(bais);
				object = ois.readObject();
				ois.close();
			}
		} catch (IOException e) {
			e.printStackTrace();
		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		}
		return object;
	}

	/**
	 * 写入byte[]
	 * 
	 * @param filename
	 * @param datas
	 */

	public static void write(String filename, byte[] datas) {
		if (datas == null || datas.length == 0) {
			return;
		}
		String filePath = ExcelUtil.currentWorkDir + filename;
		System.out.println(filePath);
		OutputStream out = null;
		try {
			out = new BufferedOutputStream(new FileOutputStream(filePath));
			out.write(datas);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (out != null) {
				try {
					out.close();
				} catch (IOException e) {
					e.printStackTrace();
				}
			}
		}
	}

	/**
	 * 创建workbook
	 * 
	 * @param in
	 * @return
	 * @throws IOException
	 * @throws InvalidFormatException
	 */

	public static Workbook createWorkbook(InputStream in) throws IOException,
			InvalidFormatException {
		if (!in.markSupported()) {
			in = new PushbackInputStream(in, 8);
		}
		if (POIFSFileSystem.hasPOIFSHeader(in)) {
			return new HSSFWorkbook(in);
		}
		if (POIXMLDocument.hasOOXMLHeader(in)) {
			return new XSSFWorkbook(OPCPackage.open(in));
		}
		throw new IllegalArgumentException("你的excel版本目前poi解析不了");
	}

	/**
	 * 列的编号转换，从字母表示到索引值
	 * 
	 * @param strbh
	 * @return
	 */
	public static int getColIndex(String strbh) {
		String straz = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
		int azlen = straz.length();
		strbh = strbh.toUpperCase();
		int col = -1;
		for (int i = 0; i < strbh.length(); i++) {
			col = col + 1;
			col = col * azlen + straz.indexOf(strbh.substring(i, i + 1));
		}
		return col;
	}

	/**
	 * 获得颜色
	 * 
	 * @param xc
	 * @return
	 */

	public static String getJavaColor(XSSFColor xc) {
		if (xc == null) {
			return "";
		}
		byte[] srgb = xc.getRgb();
		if (srgb == null) {
			return "000000";
		}
		String str1 = "";
		for (int i = 0; i < srgb.length; i++) {
			String str2 = "0" + Integer.toHexString(srgb[i]).toUpperCase();
			str1 += str2.substring(str2.length() - 2);

		}
		return str1;
	}

	/**
	 * 获得颜色
	 * 
	 * @param rgb
	 * @return
	 */

	public static String getRGB(int[] rgb) {
		int rgbr = rgb[0];
		int ggbb = rgb[1];
		int bgb = rgb[2];
		return "rgb(" + rgbr + "," + ggbb + "," + bgb + ")";
	}

	/**
	 * 获得ExcelColor对象
	 * 
	 * @param color
	 * @return
	 */

	public static ExcelColor getColor(String color) {
		color = color.substring(4, color.length() - 1);
		String[] colors = color.split(",");
		return new ExcelColor(Integer.valueOf(colors[0].trim()),
				Integer.valueOf(colors[1].trim()), Integer.valueOf(colors[2]
						.trim()));
	}

	/**
	 * 获得页面宽度
	 * 
	 * @param widthUnits
	 * @return
	 */

	public static int getPageWidth(int widthUnits) {
		int pixels = widthUnits / 256 * 7;

		int offsetWidthUnits = widthUnits % 256;

		pixels = pixels + Math.round(offsetWidthUnits / 36.57143F);

		return pixels;
	}

	public static int getPageHeight(int height) {
		return Math.round(height / 20) + 7;
	}

//	public static Length getLength(short length) {
//		return new Length(length / 20.0F, Length.UNIT.PT);
//	}
//
//	public static String getLengthStr(Length len) {
//		return len.getLength() + len.getUnit().toString();
//	}

	/**
	 * 判断是不是为2003的excel
	 * 
	 * @param is
	 * @return
	 */

	public static boolean isExcel2003(InputStream is) {
		try {
			new HSSFWorkbook(is);
		} catch (Exception e) {
			return false;
		}
		return true;
	}
}
