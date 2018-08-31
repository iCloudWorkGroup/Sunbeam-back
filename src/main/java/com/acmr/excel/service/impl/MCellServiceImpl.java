
package com.acmr.excel.service.impl;

import java.io.UnsupportedEncodingException;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MColDao;
import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.CellContent;
import com.acmr.excel.model.CellFormate;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.complete.Border;
import com.acmr.excel.model.complete.Content;
import com.acmr.excel.model.complete.CustomProp;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.complete.OperProp;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MCellService;
import com.acmr.excel.util.CellFormateUtil;

@Service("mcellService")
public class MCellServiceImpl implements MCellService {

	@Resource
	public MRowColDao mrowColDao;
	@Resource
	private BaseDao baseDao;
	@Resource
	private MSheetDao msheetDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MRowDao mrowDao;
	@Resource
	private MRowColCellDao mrowColCellDao;
	private MCol MCol;

	public void saveContent(CellContent cell, int step, String excelId) {
		String sheetId = excelId + 0;
		int rowIndex = cell.getCoordinate().getStartRow();
		int colIndex = cell.getCoordinate().getStartCol();

		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId, sheetId);
		mrowColDao.getRowList(sortRcList, excelId, sheetId);
		String rowAlias = sortRcList.get(rowIndex).getAlias();
		String colAlias = sortClList.get(colIndex).getAlias();
		String id = rowAlias + "_" + colAlias;
		MCell mcell = mcellDao.getMCell(excelId, sheetId, id);
		MSheet msheet = msheetDao.getMSheet(excelId, sheetId);
		if(null!=msheet.getProtect()&&msheet.getProtect()){
	         if((null==mcell)||(mcell.getContent().getLocked())){
	        	 return;
	         }
		}
		if (null == mcell) {
			mcell = new MCell();
			mcell.setRowspan(1);
			mcell.setColspan(1);
			mcell.setId(id);
			mcell.setSheetId(sheetId);

			MRow mrow = mrowDao.getMRow(excelId, sheetId, rowAlias);
			OperProp rp = mrow.getProps();
			MCol mcol = mcolDao.getMCol(excelId, sheetId, colAlias);
			OperProp cp = mcol.getProps();
			setMCellProperty(mcell, rp.getContent(), rp.getBorder(),
					rp.getCustomProp());
			setMCellProperty(mcell, cp.getContent(), cp.getBorder(),
					cp.getCustomProp());

			MRowColCell mrowColCell = new MRowColCell();
			mrowColCell.setRow(rowAlias);
			mrowColCell.setCol(colAlias);
			mrowColCell.setCellId(id);
			mrowColCell.setSheetId(sheetId);
			baseDao.insert(excelId, mrowColCell);// 存关系映射表
		}

		String text = cell.getContent();
		try {
			text = java.net.URLDecoder.decode(text, "utf-8");
			Content content = mcell.getContent();
			content.setType("routine");
			content.setExpress("G");
			content.setTexts(text);
			CellFormateUtil.setShowText(mcell.getContent());

			baseDao.update(excelId, mcell);// 存cell对象

			List<RowHeight> effect = cell.getEffect();
			if (null != effect) {
				for (RowHeight rh : effect) {
					String row = sortRcList.get(rh.getRow()).getAlias();
					mrowDao.updateRowHeight(excelId, sheetId, row,
							rh.getOffset());
				}
			}

			msheetDao.updateStep(excelId, sheetId, step);

		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}

	}

	@Override
	public void updateFontFamily(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String family = cell.getFamily();

		try {
			updateContent(coordinates, "family", family, null, null, null, step,
					excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateFontSize(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String size = cell.getSize();

		try {
			updateContent(coordinates, "size", size, null, null, null, step,
					excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateFontWeight(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean weight = cell.isWeight();

		try {
			updateContent(coordinates, "weight", weight, null, null, null, step,
					excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateFontItalic(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean italic = cell.getItalic();

		try {
			updateContent(coordinates, "italic", italic, null, null, null, step,
					excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateFontColor(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String color = cell.getColor();

		try {
			updateContent(coordinates, "color", color, null, null, null, step,
					excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateWordwrap(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean wordWrap = cell.getAuto();
		List<RowHeight> effect = cell.getEffect();
		try {
			updateContent(coordinates, "wordWrap", wordWrap, null, null, effect,
					step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateBgColor(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String background = cell.getColor();

		try {
			updateContent(coordinates, "background", background, null, null,
					null, step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateAlignlevel(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String alignRow = cell.getAlign();

		try {
			updateContent(coordinates, "alignRow", alignRow, null, null, null,
					step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateAlignvertical(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String alignCol = cell.getAlign();

		try {
			updateContent(coordinates, "alignCol", alignCol, null, null, null,
					step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void updateFontUnderline(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		int underline = cell.getUnderline();

		try {
			updateContent(coordinates, "underline", underline, null, null, null,
					step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {

			e.printStackTrace();
		}

	}

	@Override
	public void mergeCell(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		List<Coordinate> coordinates = cell.getCoordinate();
		List<String> colList = new ArrayList<String>();
		List<String> rowList = new ArrayList<String>();

		for (Coordinate cd : coordinates) {
			int startRow = cd.getStartRow();
			int endRow = cd.getEndRow();
			int startCol = cd.getStartCol();
			int endCol = cd.getEndCol();
			if ((endCol > -1) && (endRow > -1)) {
				for (int i = startRow; i < endRow + 1; i++) {
					rowList.add(sortRList.get(i).getAlias());
				}
				for (int i = startCol; i < endCol + 1; i++) {
					colList.add(sortCList.get(i).getAlias());
				}
			}

			if (rowList.size() > 0 && colList.size() > 0) {
				List<MRowColCell> relationList = mrowColCellDao
						.getMRowColCellList(excelId, sheetId, rowList, colList);
				mrowColCellDao.delMRowColCellList(excelId, sheetId, rowList, colList);// 删除关系表
				Map<String, String> relationMap = new HashMap<String, String>();
				for (MRowColCell mrcc : relationList) {
					relationMap.put(mrcc.getRow() + "_" + mrcc.getCol(),
							mrcc.getCellId());
				}
				int count = 0;// 找到合并区域的第一个单元格
				List<String> cellIdList = new ArrayList<String>();// 存储需要删除的MCell的id
				String id = rowList.get(0) + "_" + colList.get(0);// 合并单元格的id
				relationList.clear();// 清空关系表，用于存储新创建的对象
				MCell mc = null;// 合并单元格的MCell对象
				for (String row : rowList) {
					for (String col : colList) {
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null != cellId) {
							if (count == 0) {
								count++;
								mc = mcellDao.getMCell(excelId, sheetId,
										cellId);
								mc.setId(id);
								mc.setRowspan(rowList.size());
								mc.setColspan(colList.size());
								cellIdList.add(cellId);
							} else {
								cellIdList.add(cellId);
							}

						}

						MRowColCell mrcc = new MRowColCell();
						mrcc.setRow(row);
						mrcc.setCol(col);
						mrcc.setSheetId(sheetId);
						mrcc.setCellId(id);
						relationList.add(mrcc);
					}
				}
				if (count == 0) {
					mc = new MCell();
					mc.setId(id);
					mc.setSheetId(sheetId);
					mc.setRowspan(rowList.size());
					mc.setColspan(colList.size());
				}
				// 删除合并区域老的MCell
				mcellDao.delMCell(excelId, sheetId, cellIdList);
				// 保存新的关系表
				baseDao.insertList(excelId, relationList);
				baseDao.insert(excelId, mc);
				// 更新步骤
				msheetDao.updateStep(excelId, sheetId, step);
			}

		}
	}

	@Override
	public void splitCell(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rowMap = new HashMap<String, Integer>();
		Map<String, Integer> colMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			rowMap.put(sortRList.get(i).getAlias(), i);
		}
		for (int i = 0; i < sortCList.size(); i++) {
			colMap.put(sortCList.get(i).getAlias(), i);
		}

		List<Coordinate> coordinates = cell.getCoordinate();
		List<String> colList = new ArrayList<String>();
		List<String> rowList = new ArrayList<String>();

		for (Coordinate cd : coordinates) {
			int startRow = cd.getStartRow();
			int endRow = cd.getEndRow();
			int startCol = cd.getStartCol();
			int endCol = cd.getEndCol();
			if ((endCol > -1) && (endRow > -1)) {
				for (int i = startRow; i < endRow + 1; i++) {
					rowList.add(sortRList.get(i).getAlias());
				}
				for (int i = startCol; i < endCol + 1; i++) {
					colList.add(sortCList.get(i).getAlias());
				}
			}

			if (rowList.size() > 0 && colList.size() > 0) {
				List<MRowColCell> relationList = mrowColCellDao
						.getMRowColCellList(excelId, sheetId, rowList, colList);
				List<String> cellIdList = new ArrayList<String>();
				for (MRowColCell mrcc : relationList) {
					cellIdList.add(mrcc.getCellId());
				}
				List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
						cellIdList);
				List<MCell> mcList = new ArrayList<MCell>();// 用于存储新生成的MCell对象
				cellIdList.clear();// 清空，用于存贮是合并单元格的id
				relationList.clear();// 清空，用于存储新的关系表
				for (MCell mc : mcellList) {
					if (mc.getColspan() > 1 || mc.getRowspan() > 1) {

						cellIdList.add(mc.getId());
						String[] ids = mc.getId().split("_");
						int rowIndex = rowMap.get(ids[0]);
						int colIndex = colMap.get(ids[1]);
						for (int i = 0; i < mc.getRowspan(); i++) {
							for (int j = 0; j < mc.getColspan(); j++) {
								String row = sortRList.get(rowIndex + i)
										.getAlias();
								String col = sortCList.get(colIndex + j)
										.getAlias();
								MCell mcell = new MCell(mc);
								String id = row + "_" + col;
								mcell.setId(id);
								mcell.setColspan(1);
								mcell.setRowspan(1);
								if (!((i == 0) && (j == 0))) {
									mcell.getContent().setDisplayTexts(null);
									mcell.getContent().setTexts(null);
								}
								mcList.add(mcell);
								// 存关系表
								MRowColCell mrcc = new MRowColCell();
								mrcc.setCellId(id);
								mrcc.setRow(row);
								mrcc.setCol(col);
								mrcc.setSheetId(sheetId);
								relationList.add(mrcc);
							}
						}

					}
				}
				// 删除老的合并单元格关系表
				mrowColCellDao.delMRowColCellList1(excelId, sheetId, cellIdList);
				// 插入拆分之后的关系表
				baseDao.insertList(excelId, relationList);
				// 删除合并单元格
				mcellDao.delMCell(excelId, sheetId, cellIdList);
				// 存入拆分之后的单元格
				baseDao.insertList(excelId, mcList);
				// 更新步骤
				msheetDao.updateStep(excelId, sheetId, step);
			}

		}

	}

	@Override
	public void updateBorder(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rMap = new HashMap<String, Integer>();
		Map<String, Integer> cMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			RowCol rc = sortRList.get(i);
			rMap.put(rc.getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortCList.size(); i++) {
			RowCol rc = sortCList.get(i);
			cMap.put(rc.getAlias(), i);
		}

		List<Coordinate> coordinates = cell.getCoordinate();

		// 选中一定区域的行、列集合
		List<Object[]> blockList = new ArrayList<Object[]>();
		List<List<String>> wholeRowList = new ArrayList<List<String>>();
		List<List<String>> wholeColList = new ArrayList<List<String>>();
		for (Coordinate cd : coordinates) {
			int startRow = cd.getStartRow();
			int endRow = cd.getEndRow();
			int startCol = cd.getStartCol();
			int endCol = cd.getEndCol();
			if ((endCol > -1) && (endRow > -1)) {
				List<String> colList = new ArrayList<String>();
				List<String> rowList = new ArrayList<String>();
				Object[] list = new Object[2];
				for (int i = startRow; i < endRow + 1; i++) {
					rowList.add(sortRList.get(i).getAlias());
				}
				list[0] = rowList;
				for (int i = startCol; i < endCol + 1; i++) {
					colList.add(sortCList.get(i).getAlias());
				}
				list[1] = colList;
				blockList.add(list);
			}

			if ((endCol == -1) && (endRow > -1)) {
				List<String> row = new ArrayList<String>();
				for (int i = startRow; i < endRow + 1; i++) {
					row.add(sortRList.get(i).getAlias());
				}
				wholeRowList.add(row);
			}

			if ((endRow == -1) && (endCol > -1)) {
				List<String> col = new ArrayList<String>();
				for (int i = startCol; i < endCol + 1; i++) {
					col.add(sortCList.get(i).getAlias());
				}
				wholeColList.add(col);
			}
		}

		String type = cell.getDirection();
		int line = cell.getLine();

		List<Object> tempList = new ArrayList<Object>();// 用于存储新new的关系对象及MCell对象

		if (blockList.size() > 0) {

			for (Object[] o : blockList) {
				List<String> rowList = (List<String>) o[0];
				List<String> colList = (List<String>) o[1];
				List<MRowColCell> list1 = mrowColCellDao.getMRowColCellList(excelId,
						sheetId, rowList, colList);
				Map<String, String> relationMap = new HashMap<String, String>();
				for (MRowColCell mrcc : list1) {
					String row = mrcc.getRow();
					String col = mrcc.getCol();
					relationMap.put(row + "_" + col, mrcc.getCellId());
				}
				List<String> cellIdList = new ArrayList<String>();// 用于存储需要修改单元格的id
				switch (type) {
				case "left":
					for (String row : rowList) {
						String col = colList.get(0);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "left",
									line);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("left", line, cellIdList, excelId,
							sheetId);
					break;
				case "top":
					for (String col : colList) {
						String row = rowList.get(0);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "top",
									line);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("top", line, cellIdList, excelId,
							sheetId);
					break;
				case "right":
					for (String row : rowList) {
						String col = colList.get(colList.size() - 1);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "right",
									line);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("right", line, cellIdList, excelId,
							sheetId);
					break;
				case "bottom":
					for (String col : colList) {
						String row = rowList.get(rowList.size() - 1);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "bottom",
									line);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("bottom", line, cellIdList, excelId,
							sheetId);
					break;
				case "none":
					for (String row : rowList) {
						for (String col : colList) {
							String key = row + "_" + col;
							String cellId = relationMap.get(key);
							if (null == cellId) {
								creatMCell(sheetId, row, col, tempList, "none",
										line);
							} else {
								cellIdList.add(cellId);
							}
						}
					}
					mcellDao.updateBorder("none", line, cellIdList, excelId,
							sheetId);
					break;
				case "all":
					for (String row : rowList) {
						for (String col : colList) {
							String key = row + "_" + col;
							String cellId = relationMap.get(key);
							if (null == cellId) {
								creatMCell(sheetId, row, col, tempList, "none",
										line);
							} else {
								cellIdList.add(cellId);
							}
						}
					}
					mcellDao.updateBorder("none", line, cellIdList, excelId,
							sheetId);
					break;
				case "outer":
					cellIdList.clear();
					for (String row : rowList) {
						String col = colList.get(0);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "left",
									line);
							baseDao.insertList(excelId, tempList);// 存储新创建的关系表及MCell对象
							tempList.clear();
							relationMap.put(key, key);
							
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("left", line, cellIdList, excelId,
							sheetId);
					cellIdList.clear();
					for (String col : colList) {
						String row = rowList.get(0);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "top",
									line);
							baseDao.insertList(excelId, tempList);// 存储新创建的关系表及MCell对象
							tempList.clear();
							relationMap.put(key, key);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("top", line, cellIdList, excelId,
							sheetId);
					cellIdList.clear();
					for (String row : rowList) {
						String col = colList.get(colList.size() - 1);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "right",
									line);
							baseDao.insertList(excelId, tempList);// 存储新创建的关系表及MCell对象
							tempList.clear();
							relationMap.put(key, key);
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("right", line, cellIdList, excelId,
							sheetId);
					cellIdList.clear();
					for (String col : colList) {
						String row = rowList.get(rowList.size() - 1);
						String key = row + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							creatMCell(sheetId, row, col, tempList, "bottom",
									line);
							//relationMap.put(key, key);最后一个不用加
						} else {
							cellIdList.add(cellId);
						}
					}
					mcellDao.updateBorder("bottom", line, cellIdList, excelId,
							sheetId);
					break;
				default:
					break;
				}
			}
		}

		if (wholeRowList.size() > 0) {
			for (List<String> rowList : wholeRowList) {
				switch (type) {
				case "top":
					String row = rowList.get(0);
					List<MRowColCell> relationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, row, "row");
					List<String> cellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : relationList) {
						cellIdList.add(mrcc.getCellId());
					}
					List<MCell> mcellList = mcellDao.getMCellList(excelId,
							sheetId, cellIdList);
					cellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : mcellList) {
						if (mc.getRowspan() == 1) {
							cellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							if ((row.equals(ids[0]))
									&& (mc.getRowspan() < rowList.size() + 1)) {
								cellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("top", line, cellIdList, excelId,
							sheetId);
					mrowDao.updateBorder("top", line, row, excelId, sheetId);
					break;
				case "bottom":
					String bottomRow = rowList.get(rowList.size() - 1);
					List<MRowColCell> bottomRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, bottomRow,
									"row");
					List<String> bottomCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : bottomRelationList) {
						bottomCellIdList.add(mrcc.getCellId());
					}
					List<MCell> bottomMcellList = mcellDao.getMCellList(excelId,
							sheetId, bottomCellIdList);
					bottomCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : bottomMcellList) {
						if (mc.getRowspan() == 1) {
							bottomCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int index = rMap.get(ids[0]);
							String row1 = sortRList
									.get(index + mc.getRowspan() - 1)
									.getAlias();
							if ((bottomRow.equals(row1))
									&& (mc.getRowspan() < rowList.size() + 1)) {
								bottomCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("bottom", line, bottomCellIdList,
							excelId, sheetId);
					mrowDao.updateBorder("bottom", line, bottomRow, excelId,
							sheetId);
					break;
				case "none":
					List<MRowColCell> noneRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, rowList,
									"row");
					List<String> noneCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : noneRelationList) {
						noneCellIdList.add(mrcc.getCellId());
					}
					List<MCell> noneMcellList = mcellDao.getMCellList(excelId,
							sheetId, noneCellIdList);
					noneCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : noneMcellList) {
						if (mc.getRowspan() == 1) {
							noneCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int begin = rMap.get(ids[0]);
							int end = begin + mc.getRowspan() - 1;
							int rowBegin = rMap.get(rowList.get(0));
							int rowEnd = rMap
									.get(rowList.get(rowList.size() - 1));

							if ((begin >= rowBegin) && (end <= rowEnd)) {
								noneCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("none", line, noneCellIdList, excelId,
							sheetId);
					mrowDao.updateBorderList("none", line, rowList, excelId,
							sheetId);
					break;
				case "all":
					List<MRowColCell> allRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, rowList,
									"row");
					List<String> allCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : allRelationList) {
						allCellIdList.add(mrcc.getCellId());
					}
					List<MCell> allMcellList = mcellDao.getMCellList(excelId,
							sheetId, allCellIdList);
					allCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : allMcellList) {
						if (mc.getRowspan() == 1) {
							allCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int begin = rMap.get(ids[0]);
							int end = begin + mc.getRowspan() - 1;
							int rowBegin = rMap.get(rowList.get(0));
							int rowEnd = rMap
									.get(rowList.get(rowList.size() - 1));

							if ((begin >= rowBegin) && (end <= rowEnd)) {
								allCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("all", line, allCellIdList, excelId,
							sheetId);
					mrowDao.updateBorderList("all", line, rowList, excelId,
							sheetId);
					break;
				case "outer":
					String outerRow = rowList.get(0);
					List<MRowColCell> outerRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, outerRow,
									"row");
					List<String> outerCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : outerRelationList) {
						outerCellIdList.add(mrcc.getCellId());
					}
					List<MCell> ourerMcellList = mcellDao.getMCellList(excelId,
							sheetId, outerCellIdList);
					outerCellIdList.clear();// 清空后，存储符合条件的单元格
					for (MCell mc : ourerMcellList) {
						if (mc.getRowspan() == 1) {
							outerCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							if ((outerRow.equals(ids[0]))
									&& (mc.getRowspan() < rowList.size() + 1)) {
								outerCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("top", line, outerCellIdList, excelId,
							sheetId);
					mrowDao.updateBorder("top", line, outerRow, excelId,
							sheetId);

					String outerRow1 = rowList.get(rowList.size() - 1);
					List<MRowColCell> outerRelationList1 = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, outerRow1,
									"row");
					List<String> outerCellIdList1 = new ArrayList<String>();
					for (MRowColCell mrcc : outerRelationList1) {
						outerCellIdList1.add(mrcc.getCellId());
					}
					List<MCell> outerMcellList1 = mcellDao.getMCellList(excelId,
							sheetId, outerCellIdList1);
					outerCellIdList1.clear();// 清空存储符合条件的单元格
					for (MCell mc : outerMcellList1) {
						if (mc.getRowspan() == 1) {
							outerCellIdList1.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int index = rMap.get(ids[0]);
							String row1 = sortRList
									.get(index + mc.getRowspan() - 1)
									.getAlias();
							if ((outerRow1.equals(row1))
									&& (mc.getRowspan() < rowList.size() + 1)) {
								outerCellIdList1.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("bottom", line, outerCellIdList1,
							excelId, sheetId);
					mrowDao.updateBorder("bottom", line, outerRow1, excelId,
							sheetId);
					break;
				default:
					break;
				}
			}
		}

		if (wholeColList.size() > 0) {
			for (List<String> colList : wholeColList) {
				switch (type) {
				case "left":
					String col = colList.get(0);
					List<MRowColCell> relationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, col, "col");
					List<String> cellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : relationList) {
						cellIdList.add(mrcc.getCellId());
					}
					List<MCell> mcellList = mcellDao.getMCellList(excelId,
							sheetId, cellIdList);
					cellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : mcellList) {
						if (mc.getColspan() == 1) {
							cellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							if ((col.equals(ids[1]))
									&& (mc.getColspan() < colList.size() + 1)) {
								cellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("left", line, cellIdList, excelId,
							sheetId);
					mcolDao.updateBorder("left", line, col, excelId, sheetId);
					break;
				case "right":
					String rightCol = colList.get(colList.size() - 1);
					List<MRowColCell> rightRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, rightCol,
									"col");
					List<String> rightCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : rightRelationList) {
						rightCellIdList.add(mrcc.getCellId());
					}
					List<MCell> rightMcellList = mcellDao.getMCellList(excelId,
							sheetId, rightCellIdList);
					rightCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : rightMcellList) {
						if (mc.getRowspan() == 1) {
							rightCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int index = rMap.get(ids[1]);
							String col1 = sortCList
									.get(index + mc.getColspan() - 1)
									.getAlias();
							if ((rightCol.equals(col1))
									&& (mc.getColspan() < colList.size() + 1)) {
								rightCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("right", line, rightCellIdList,
							excelId, sheetId);
					mcolDao.updateBorder("right", line, rightCol, excelId,
							sheetId);
					break;
				case "none":
					List<MRowColCell> noneRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, colList,
									"col");
					List<String> noneCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : noneRelationList) {
						noneCellIdList.add(mrcc.getCellId());
					}
					List<MCell> noneMcellList = mcellDao.getMCellList(excelId,
							sheetId, noneCellIdList);
					noneCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : noneMcellList) {
						if (mc.getColspan() == 1) {
							noneCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int begin = cMap.get(ids[1]);
							int end = begin + mc.getColspan() - 1;
							int colBegin = cMap.get(colList.get(0));
							int colEnd = cMap
									.get(colList.get(colList.size() - 1));

							if ((begin >= colBegin) && (end <= colEnd)) {
								noneCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("none", line, noneCellIdList, excelId,
							sheetId);
					mcolDao.updateBorderList("none", line, colList, excelId,
							sheetId);
					break;
				case "all":
					List<MRowColCell> allRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, colList,
									"col");
					List<String> allCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : allRelationList) {
						allCellIdList.add(mrcc.getCellId());
					}
					List<MCell> allMcellList = mcellDao.getMCellList(excelId,
							sheetId, allCellIdList);
					allCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : allMcellList) {
						if (mc.getColspan() == 1) {
							allCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int begin = cMap.get(ids[1]);
							int end = begin + mc.getColspan() - 1;
							int rowBegin = cMap.get(colList.get(0));
							int rowEnd = cMap
									.get(colList.get(colList.size() - 1));

							if ((begin >= rowBegin) && (end <= rowEnd)) {
								allCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("all", line, allCellIdList, excelId,
							sheetId);
					mcolDao.updateBorderList("all", line, colList, excelId,
							sheetId);
					break;
				case "outer":
					String outerCol = colList.get(0);
					List<MRowColCell> outerRelationList = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, outerCol,
									"col");
					List<String> outerCellIdList = new ArrayList<String>();
					for (MRowColCell mrcc : outerRelationList) {
						outerCellIdList.add(mrcc.getCellId());
					}
					List<MCell> ourerMcellList = mcellDao.getMCellList(excelId,
							sheetId, outerCellIdList);
					outerCellIdList.clear();// 清空存储符合条件的单元格
					for (MCell mc : ourerMcellList) {
						if (mc.getColspan() == 1) {
							outerCellIdList.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							if ((outerCol.equals(ids[1]))
									&& (mc.getRowspan() < colList.size() + 1)) {
								outerCellIdList.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("left", line, outerCellIdList,
							excelId, sheetId);
					mcolDao.updateBorder("left", line, outerCol, excelId,
							sheetId);

					String outerCol1 = colList.get(colList.size() - 1);
					List<MRowColCell> outerRelationList1 = mrowColCellDao
							.getMRowColCellList(excelId, sheetId, outerCol1,
									"col");
					List<String> outerCellIdList1 = new ArrayList<String>();
					for (MRowColCell mrcc : outerRelationList1) {
						outerCellIdList1.add(mrcc.getCellId());
					}
					List<MCell> outerMcellList1 = mcellDao.getMCellList(excelId,
							sheetId, outerCellIdList1);
					outerCellIdList1.clear();// 清空存储符合条件的单元格
					for (MCell mc : outerMcellList1) {
						if (mc.getColspan() == 1) {
							outerCellIdList1.add(mc.getId());
						} else {
							String[] ids = mc.getId().split("_");
							int index = cMap.get(ids[1]);
							String col1 = sortCList
									.get(index + mc.getColspan() - 1)
									.getAlias();
							if ((outerCol1.equals(col1))
									&& (mc.getRowspan() < colList.size() + 1)) {
								outerCellIdList1.add(mc.getId());
							}
						}
					}
					mcellDao.updateBorder("right", line, outerCellIdList1,
							excelId, sheetId);
					mcolDao.updateBorder("right", line, outerCol1, excelId,
							sheetId);
					break;
				default:
					break;
				}
			}
		}
		// 更新MCell对象属性
		baseDao.insertList(excelId, tempList);// 存储新创建的关系表及MCell对象
		msheetDao.updateStep(excelId, sheetId, step);

	}

	@Override
	public void updateFormat(CellFormate cellFormate, int step,
			String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cellFormate.getCoordinate();
		String value1 = cellFormate.getFormat();
		String value2 = cellFormate.getExpress();

		try {
			updateContent(coordinates, "type", value1, "express", value2, null,
					step, excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException
				| IllegalArgumentException | IllegalAccessException e) {
			e.printStackTrace();
		}

	}

	private void creatMCell(String sheetId, String row, String col,
			List<Object> tempList, String property, Object value) {

		MRowColCell mr = new MRowColCell();
		mr.setCellId(row + "_" + col);
		mr.setRow(row);
		mr.setCol(col);
		mr.setSheetId(sheetId);
		tempList.add(mr);
		MCell mc = new MCell(row + "_" + col, sheetId);
		Border border = mc.getBorder();
		if ("none".equals(property)) {
			border.setLeft((int) value);
			border.setTop((int) value);
			border.setRight((int) value);
			border.setBottom((int) value);
		} else if ("all".equals(property)) {
			border.setLeft((int) value);
			border.setTop((int) value);
			border.setRight((int) value);
			border.setBottom((int) value);
		} else {
			Field f;
			try {
				f = border.getClass().getDeclaredField(property);
				f.setAccessible(true);
				f.set(mc.getBorder(), value);
			} catch (Exception e) {
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}
		tempList.add(mc);
	}

	public void updateContent(List<Coordinate> coordinates, String property1,
			Object value1, String property2, Object value2,
			List<RowHeight> effect, int step, String excelId, String sheetId)
					throws NoSuchFieldException, SecurityException,
					IllegalArgumentException, IllegalAccessException {

		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String, Integer> rMap = new HashMap<String, Integer>();
		Map<String, Integer> cMap = new HashMap<String, Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for (int i = 0; i < sortRList.size(); i++) {
			RowCol rc = sortRList.get(i);
			rMap.put(rc.getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for (int i = 0; i < sortCList.size(); i++) {
			RowCol rc = sortCList.get(i);
			cMap.put(rc.getAlias(), i);
		}

		List<String> colList = new ArrayList<String>();
		List<String> rowList = new ArrayList<String>();
		List<List<String>> wholeRowList = new ArrayList<List<String>>();
		List<List<String>> wholeColList = new ArrayList<List<String>>();
		for (Coordinate cd : coordinates) {
			int startRow = cd.getStartRow();
			int endRow = cd.getEndRow();
			int startCol = cd.getStartCol();
			int endCol = cd.getEndCol();
			if ((endCol > -1) && (endRow > -1)) {
				for (int i = startRow; i < endRow + 1; i++) {
					rowList.add(sortRList.get(i).getAlias());
				}
				for (int i = startCol; i < endCol + 1; i++) {
					colList.add(sortCList.get(i).getAlias());
				}
			}

			if ((endCol == -1) && (endRow > -1)) {
				List<String> row = new ArrayList<String>();
				for (int i = startRow; i < endRow + 1; i++) {
					row.add(sortRList.get(i).getAlias());
				}
				wholeRowList.add(row);
			}

			if ((endRow == -1) && (endCol > -1)) {
				List<String> col = new ArrayList<String>();
				for (int i = startCol; i < endCol + 1; i++) {
					col.add(sortCList.get(i).getAlias());
				}
				wholeColList.add(col);
			}
		}

		List<Object> tempList = new ArrayList<Object>();// 用于存储新new的关系对象及MCell对象
		List<String> cellIdList = new ArrayList<String>();// 用于存储需要修改单元格的id
		if (rowList.size() > 0 && colList.size() > 0) {
			List<MRowColCell> list1 = mrowColCellDao.getMRowColCellList(excelId,
					sheetId, rowList, colList);
			Map<String, String> relationMap = new HashMap<String, String>();
			for (MRowColCell mrcc : list1) {
				String row = mrcc.getRow();
				String col = mrcc.getCol();
				relationMap.put(row + "_" + col, mrcc.getCellId());
			}
			for (Coordinate c : coordinates) {
				int startRow = c.getStartRow();
				int endRow = c.getEndRow();
				int startCol = c.getStartCol();
				int endCol = c.getEndCol();
				if (endRow > -1 && endCol > -1) {
					for (int i = startRow; i < endRow + 1; i++) {
						String row = sortRList.get(i).getAlias();
						for (int j = startCol; j < endCol + 1; j++) {
							String col = sortCList.get(j).getAlias();
							String cellId = relationMap.get(row + "_" + col);
							if (null == cellId) {
								MRowColCell mr = new MRowColCell();
								mr.setCellId(row + "_" + col);
								mr.setRow(row);
								mr.setCol(col);
								mr.setSheetId(sheetId);
								tempList.add(mr);
								MCell mc = new MCell(row + "_" + col, sheetId);
								
								
								 // 将行、列自带的属性赋值给单元格 
								 MRow mrow=mrowDao.getMRow(excelId, sheetId, row);
								 OperProp rp = mrow.getProps(); 
								 MCol mcol= mcolDao.getMCol(excelId, sheetId, col);
								 OperProp cp = mcol.getProps();
								 setMCellProperty(mc, rp.getContent(),
								 rp.getBorder(), rp.getCustomProp());
								 setMCellProperty(mc, cp.getContent(),
								 cp.getBorder(), cp.getCustomProp());
								 
								
								if ("type".equals(property1)
										&& "express".equals(property2)) {
									Field f = mc.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mc.getContent(), value1);
									Field f2 = mc.getContent().getClass()
											.getDeclaredField(property2);
									f2.setAccessible(true);
									f2.set(mc.getContent(), value2);
								} else {
									Field f = mc.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mc.getContent(), value1);
								}
								
								
								
								tempList.add(mc);
							} else {
								// 当时设置单元格数字格式时，需要修改displayTexts值
								if ("type".equals(property1)) {
									MCell mc = mcellDao.getMCell(excelId,
											sheetId, cellId);
									Content content = mc.getContent();
									content.setType(value1.toString());
									content.setExpress(value2.toString());
									CellFormateUtil.setShowText(content);
									baseDao.update(excelId, mc);
								} else {
									cellIdList.add(cellId);
								}
							}

						}
					}
				}
			}
		}

		if (wholeRowList.size() > 0) {
			for (List<String> rows : wholeRowList) {
				List<MRowColCell> list2 = mrowColCellDao.getMRowColCellList(excelId,
						sheetId, rows, "row");
				Map<String, String> relationMap = new HashMap<String, String>();
				List<String> idList = new ArrayList<String>();
				for (MRowColCell mrcc : list2) {
					String row = mrcc.getRow();
					String col = mrcc.getCol();
					relationMap.put(row + "_" + col, mrcc.getCellId());
				}
				// 获取所有的列样式对象
				List<MCol> mcolList = mcolDao.getMColList(excelId, sheetId);
				// 用存贮选中的行
				// Map<String, String> wholeRowMap = new HashMap<String,
				// String>();
				for (String row : rows) {
					// wholeRowMap.put(row, row);
					for (MCol mc : mcolList) {
						String key = row + "_" + mc.getAlias();
						String cellId = relationMap.get(key);
						if (null == cellId) {
							Content con = mc.getProps().getContent();
							Field fCol = con.getClass()
									.getDeclaredField(property1);
							fCol.setAccessible(true);
							Object value = fCol.get(con);
							if (null != value) {
								MRowColCell mr = new MRowColCell();
								mr.setCellId(key);
								mr.setRow(row);
								mr.setCol(mc.getAlias());
								mr.setSheetId(sheetId);
								tempList.add(mr);// 关系表
								MCell mcell = new MCell(key, sheetId);
								
								// 将行、列自带的属性赋值给单元格
								MRow mrow = mrowDao.getMRow(excelId, sheetId, row);
								OperProp rp = mrow.getProps();
								OperProp cp = mc.getProps();
								setMCellProperty(mcell, rp.getContent(),
										rp.getBorder(), rp.getCustomProp());
								setMCellProperty(mcell, cp.getContent(),
										cp.getBorder(), cp.getCustomProp());
								
								
								
								if ("type".equals(property1)
										&& "express".equals(property2)) {
									Field f = mcell.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mcell.getContent(), value1);
									Field f2 = mcell.getContent().getClass()
											.getDeclaredField(property2);
									f2.setAccessible(true);
									f2.set(mcell.getContent(), value2);
								} else {
									Field f = mcell.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mcell.getContent(), value1);
								}
								
							}
						} else {
							if (!idList.contains(cellId)) {
								idList.add(cellId);
							}
						}
					}
				}
				// 找出整列操作中符合条件的合并单元格
				List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
						idList);
				// getOneCellList(mcellList, rMap, cMap, sortRList, sortCList);
				int rowEnd = rMap.get(rows.get(rows.size() - 1));
				int rowBegin = rMap.get(rows.get(0));
				for (MCell mc : mcellList) {
					String id = mc.getId();
					String[] ids = id.split("_");
					int rowIndex = rMap.get(ids[0]);
					if ((rowIndex >= rowBegin)
							&& (rowIndex + mc.getRowspan() <= rowEnd + 1)) {
						if ("type".equals(property1)) {
							// 当时设置单元格数字格式时，需要修改displayTexts值
							Content content = mc.getContent();
							content.setType(value1.toString());
							content.setExpress(value2.toString());
							CellFormateUtil.setShowText(content);
							baseDao.update(excelId, mc);
							idList.remove(mc.getId());// type和express已更新，不需重复更新
						}
					} else {
						idList.remove(mc.getId()); // 剔除不符合条件的合并单元格
					}

				}
				cellIdList.addAll(idList);
				// 更新列对象
				mrowDao.updateContent(property1, value1, property2, value2,
						rows, excelId, sheetId);
			}
		}

		if (wholeColList.size() > 0) {
			for (List<String> cols : wholeColList) {
				List<MRowColCell> list3 = mrowColCellDao.getMRowColCellList(excelId,
						sheetId, cols, "col");
				Map<String, String> relationMap = new HashMap<String, String>();
				List<String> idList = new ArrayList<String>();
				for (MRowColCell mrcc : list3) {
					String row = mrcc.getRow();
					String col = mrcc.getCol();
					relationMap.put(row + "_" + col, mrcc.getCellId());
				}
				// 获取所有的行样式对象
				List<MRow> mrowList = mrowDao.getMRowList(excelId, sheetId);
				// 存储选中列
				// Map<String, String> wholeColMap = new HashMap<String,
				// String>();
				for (String col : cols) {
					// wholeColMap.put(col, col);
					for (MRow mr : mrowList) {
						String key = mr.getAlias() + "_" + col;
						String cellId = relationMap.get(key);
						if (null == cellId) {
							Content con = mr.getProps().getContent();
							Field fRow = con.getClass()
									.getDeclaredField(property1);
							fRow.setAccessible(true);
							Object value = fRow.get(con);
							if (null != value) {
								MRowColCell mrcc = new MRowColCell();
								mrcc.setCellId(key);
								mrcc.setRow(mr.getAlias());
								mrcc.setCol(col);
								mrcc.setSheetId(sheetId);
								tempList.add(mrcc);// 关系表
								MCell mcell = new MCell(key, sheetId);
								
								// 将行、列自带的属性赋值给单元格
								OperProp rp = mr.getProps();
								MCol mcol = mcolDao.getMCol(excelId, sheetId, col);
								OperProp cp = mcol.getProps();
								setMCellProperty(mcell, rp.getContent(),
										rp.getBorder(), rp.getCustomProp());
								setMCellProperty(mcell, cp.getContent(),
										cp.getBorder(), cp.getCustomProp());
								
								
								
								if ("type".equals(property1)
										&& "express".equals(property2)) {
									Field f = mcell.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mcell.getContent(), value1);
									Field f2 = mcell.getContent().getClass()
											.getDeclaredField(property2);
									f2.setAccessible(true);
									f2.set(mcell.getContent(), value2);
								} else {
									Field f = mcell.getContent().getClass()
											.getDeclaredField(property1);
									f.setAccessible(true);
									f.set(mcell.getContent(), value1);
								}
								tempList.add(mcell);
							}
						} else {
							if (!idList.contains(cellId)) {
								idList.add(cellId);
							}
						}
					}
				}
				// 找出整列操作中部分条件的合并单元格
				List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId,
						idList);
				int colEnd = cMap.get(cols.get(cols.size() - 1));
				int colBegin = cMap.get(cols.get(0));
				for (MCell mc : mcellList) {
					String id = mc.getId();
					String[] ids = id.split("_");
					int colIndex = cMap.get(ids[1]);
					if ((colIndex >= colBegin)
							&& (colIndex + mc.getColspan() <= colEnd + 1)) {
						if ("type".equals(property1)) {
							// 当时设置单元格数字格式时，需要修改displayTexts值
							Content content = mc.getContent();
							content.setType(value1.toString());
							content.setExpress(value2.toString());
							CellFormateUtil.setShowText(content);
							baseDao.update(excelId, mc);
							idList.remove(mc.getId());// type和express已更新，不需重复更新
						}
					} else {
						idList.remove(mc.getId()); // 剔除不符合条件的合并单元格
					}

				}

				cellIdList.addAll(idList);
				// 更新列对象
				mcolDao.updateContent(property1, value1, property2, value2,
						cols, excelId, sheetId);
			}
		}
		// 更新MCell对象属性
		mcellDao.updateContent(property1, value1, property2, value2, cellIdList,
				excelId, sheetId);
		baseDao.insertList(excelId, tempList);// 存储新创建的关系表及MCell对象

		// 更新自动换行时，对行高的影响
		if (null != effect) {
			for (RowHeight rh : effect) {
				String row = sortRList.get(rh.getRow()).getAlias();
				mrowDao.updateRowHeight(excelId, sheetId, row, rh.getOffset());
			}
		}

		msheetDao.updateStep(excelId, sheetId, step);

	}

	private void getOneCellList(List<MCell> cellList, Map<String, Integer> rMap,
			Map<String, Integer> cMap, List<RowCol> sortRList,
			List<RowCol> sortCList) {
		for (MCell mc : cellList) {

			if (mc.getRowspan() == 1 && mc.getColspan() == 1) {

			} else {
				String[] ids = mc.getId().split("_");
				Occupy oc = mc.getOccupy();
				int rwoIndex = rMap.get(ids[0]);
				for (int i = 0; i < mc.getRowspan(); i++) {

					oc.getRow().add(sortRList.get(rwoIndex).getAlias());
					rwoIndex++;
				}
				int colIndex = cMap.get(ids[1]);
				for (int i = 0; i < mc.getColspan(); i++) {

					oc.getCol().add(sortCList.get(colIndex).getAlias());
					colIndex++;
				}
			}

		}
	}

	private void setMCellProperty(MCell mcell, Content content, Border border,
			CustomProp customProp) {
		try {
			Class<? extends Content> c = content.getClass();
			Content mc = mcell.getContent();
			Field[] cFields = c.getDeclaredFields();
			for (Field f : cFields) {
				f.setAccessible(true);
				Object value = f.get(content);
				if (null != value) {
					Field mf = mc.getClass().getDeclaredField(f.getName());
					mf.setAccessible(true);
					mf.set(mc, value);
				}
			}

			Class<? extends Border> b = border.getClass();
			Border mb = mcell.getBorder();
			Field[] bFields = b.getDeclaredFields();
			for (Field f : bFields) {
				f.setAccessible(true);
				Object value = f.get(border);
				if (null != value) {
					Field bf = mb.getClass().getDeclaredField(f.getName());
					bf.setAccessible(true);
					bf.set(mb, value);
				}
			}

			Class<? extends CustomProp> p = customProp.getClass();
			CustomProp mp = mcell.getCustomProp();
			Field[] pFields = p.getDeclaredFields();
			for (Field f : pFields) {
				f.setAccessible(true);
				Object value = f.get(customProp);
				if (null != value) {
					Field pf = mp.getClass().getDeclaredField(f.getName());
					pf.setAccessible(true);
					pf.set(mp, value);
				}
			}

		} catch (Exception e) {

			e.printStackTrace();
		}
	}

	@Override
	public void updateComment(Cell cell, int step, String excelId) {
		String sheetId = excelId + 0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String comment = cell.getComment();
		try {
			updateProperty(coordinates, excelId, sheetId, "comment", comment);
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
	
	@Override
	public void updateLock(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		boolean locked = cell.isLock();
		
		try {
			updateContent(coordinates, "locked", locked, null, null, null, step,
					excelId, sheetId);
		} catch (Exception e) {
			// TODO: handle exception
		}
		
		
	}

	public void updateProperty(List<Coordinate> coordinates, String excelId,
			String sheetId, String property, Object value) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		List<String> colList = new ArrayList<String>();
		List<String> rowList = new ArrayList<String>();
		// 用于存储新创建的MCell和关系对象
		List<Object> tempList = new ArrayList<Object>();
		// 存储需要更新的MCell的id
		List<String> cellIdList = new ArrayList<String>();
		for (Coordinate cd : coordinates) {
			int startRow = cd.getStartRow();
			int endRow = cd.getEndRow();
			int startCol = cd.getStartCol();
			int endCol = cd.getEndCol();
			if ((endCol > -1) && (endRow > -1)) {
				for (int i = startRow; i < endRow + 1; i++) {
					rowList.add(sortRList.get(i).getAlias());
				}
				for (int i = startCol; i < endCol + 1; i++) {
					colList.add(sortCList.get(i).getAlias());
				}
			}

			if (rowList.size() > 0 && colList.size() > 0) {
				List<MRowColCell> relationList = mrowColCellDao
						.getMRowColCellList(excelId, sheetId, rowList, colList);
				Map<String, String> relationMap = new HashMap<String, String>();
				for (MRowColCell mrcc : relationList) {
					relationMap.put(mrcc.getRow() + "_" + mrcc.getCol(),
							mrcc.getCellId());
				}
				for (int i = startRow; i < endRow + 1; i++) {
					String row = sortRList.get(i).getAlias();
					for (int j = startCol; j < endCol + 1; j++) {
						String col = sortCList.get(j).getAlias();
						String cellId = relationMap.get(row + "_" + col);
						if (null == cellId) {
							MRowColCell mr = new MRowColCell();
							mr.setCellId(row + "_" + col);
							mr.setRow(row);
							mr.setCol(col);
							mr.setSheetId(sheetId);
							tempList.add(mr);
							MCell mc = new MCell(row + "_" + col, sheetId);
							Field f = mc.getContent().getClass()
									.getDeclaredField(property);
							f.setAccessible(true);
							f.set(mc.getContent(), value);

							// 将行、列自带的属性赋值给单元格
							MRow mrow = mrowDao.getMRow(excelId, sheetId, row);
							OperProp rp = mrow.getProps();
							MCol mcol = mcolDao.getMCol(excelId, sheetId, col);
							OperProp cp = mcol.getProps();
							setMCellProperty(mc, rp.getContent(),
									rp.getBorder(), rp.getCustomProp());
							setMCellProperty(mc, cp.getContent(),
									cp.getBorder(), cp.getCustomProp());

							tempList.add(mc);
						} else {
							cellIdList.add(cellId);
						}
					}
				}
			}
		}
	   if(tempList.size()>0){
		   baseDao.insertList(excelId, tempList);
	   }
	   if(cellIdList.size()>0){
		   mcellDao.updateCustomProp(property, value, cellIdList, excelId, sheetId);
	   }
	}

	

	/*
	 * public void OperationCellByCord(ArrayList<Object> al){ //查找关系表
	 * List<MRowColCell> relationList =
	 * mcellDao.getMRowColCellList(excelId,sheetId,alias,"col"); List<String>
	 * cellIdList = new ArrayList<String>(); //Map<String,String> relationMap =
	 * new HashMap<String,String>(); for(MRowColCell mrcc:relationList){
	 * 
	 * //relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
	 * cellIdList.add(mrcc.getCellId()); } List<MCell> cellList =
	 * mcellDao.getMCellList(excelId,sheetId, cellIdList); //删除关系表
	 * mcellDao.delMRowColCell(excelId,sheetId,"col", alias);
	 * cellIdList.clear();//存需要删除的MExcelCell的Id
	 * cellList.clear();//存需要插入的MExcelCell对象
	 * 
	 * for(MCell mc:cellList){ if(mc.getColspan()==1){
	 * cellIdList.add(mc.getId()); }else{ String[] ids = mc.getId().split("_");
	 * if(ids[1].equals(alias)){ //删除老的MExcelCell cellIdList.add(mc.getId());
	 * MCell mec = mc; String id = ids[0]+"_"+backAlias; //修改合并单元格其他关系表的cellId
	 * mcellDao.updateMRowColCell(excelId,sheetId, mec.getId(), id);
	 * mec.setId(id); mec.setColspan(mc.getColspan()-1);
	 * 
	 * cellList.add(mec);//插入新MCell }else{ MCell mec = mc;
	 * mec.setColspan(mc.getColspan()-1); baseDao.update(excelId, mec); } } }
	 * 
	 * mcellDao.delMCell(excelId,sheetId, cellIdList); if(cellList.size()>0){
	 * baseDao.insert(excelId, cellList);
	 * 
	 * }
	 * 
	 * }
	 */

}
