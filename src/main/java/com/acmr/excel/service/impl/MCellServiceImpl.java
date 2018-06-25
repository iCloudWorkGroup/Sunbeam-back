
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
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.CellContent;
import com.acmr.excel.model.ConverCell;
import com.acmr.excel.model.Coordinate;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.complete.Occupy;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.service.MCellService;
import com.acmr.excel.util.CellFormateUtil;

@Service("mcellService")
public class MCellServiceImpl implements MCellService {
	
	@Resource
	private MRowColDao mrowColDao;
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

	
	public void saveContent(CellContent cell, int step, String excelId) {
		String sheetId = excelId+0;
		int rowIndex = cell.getCoordinate().getStartRow();
		int colIndex = cell.getCoordinate().getStartCol();
		/*int rowEnd = cell.getCoordinate().getEndRow();
		int colEnd = cell.getCoordinate().getEndCol();*/
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId,sheetId);
		mrowColDao.getRowList(sortRcList, excelId,sheetId);
		String rowAlias = sortRcList.get(rowIndex).getAlias();
		String colAlias = sortClList.get(colIndex).getAlias();
		String id = rowAlias+"_"+colAlias;
		MCell mcell = mcellDao.getMCell(excelId,sheetId, id);
		
		if(null == mcell){
			mcell = new MCell();
			mcell.setRowspan(1);
			mcell.setColspan(1);
			mcell.setId(id);
			mcell.setSheetId(sheetId);
			
			MRowColCell mrowColCell = new MRowColCell();
			mrowColCell.setRow(rowAlias);
			mrowColCell.setCol(colAlias);
			mrowColCell.setCellId(id);
			mrowColCell.setSheetId(sheetId);
			baseDao.insert(excelId, mrowColCell);//存关系映射表
		}
	
		String content = cell.getContent();
		try {
			content = java.net.URLDecoder.decode(content, "utf-8");
			
			mcell.getContent().setTexts(content);
			mcell.getContent().setDisplayTexts(content);
			CellFormateUtil.autoRecognise(content, mcell);
			
			
			baseDao.update(excelId, mcell);//存cell对象
				
			msheetDao.updateStep(excelId,sheetId, step);
			
		} catch (UnsupportedEncodingException e) {
			e.printStackTrace();
		}
		

	}


	@Override
	public void updateFontFamily(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String family = cell.getFamily();
		
		try {
			updateContent(coordinates, "family",  family, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}
	
	@Override
	public void updateFontSize(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String size = cell.getSize();
		
		try {
			updateContent(coordinates, "size",  size, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
	
			e.printStackTrace();
		}
		
	}
	
	@Override
	public void updateFontWeight(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean weight = cell.isWeight();
		
		try {
			updateContent(coordinates, "weight",  weight, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateFontItalic(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean italic = cell.getItalic();
		
		try {
			updateContent(coordinates, "italic",  italic, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateFontColor(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String color = cell.getColor();
		
		try {
			updateContent(coordinates, "color",  color, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateWordwrap(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		Boolean wordWrap = cell.getAuto();
		
		try {
			updateContent(coordinates, "wordWrap",  wordWrap, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateBgColor(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String background = cell.getColor();
		
		try {
			updateContent(coordinates, "background",  background, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateAlignlevel(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String alignRow = cell.getAlign();
		
		try {
			updateContent(coordinates, "alignRow",  alignRow, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}


	@Override
	public void updateAlignvertical(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<Coordinate> coordinates = cell.getCoordinate();
		String alignCol = cell.getAlign();
		
		try {
			updateContent(coordinates, "alignCol",  alignCol, step,  excelId, sheetId);
		} catch (NoSuchFieldException | SecurityException | IllegalArgumentException | IllegalAccessException e) {
			
			e.printStackTrace();
		}
		
	}
	
	@Override
	public void mergeCell(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		List<Coordinate> coordinates = cell.getCoordinate();
		   List<String> colList = new ArrayList<String>();
	       List<String> rowList = new ArrayList<String>();
	       
	        for(Coordinate cd:coordinates){
	        	int startRow = cd.getStartRow();
	        	int endRow = cd.getEndRow();
	        	int startCol = cd.getStartCol();
	        	int endCol = cd.getEndCol();
	        	if((endCol> -1)&&(endRow > -1)){
		        	for(int i = startRow;i<endRow+1;i++){
		        		rowList.add(sortRList.get(i).getAlias());
		        	}
		        	for(int i = startCol;i<endCol+1;i++){
		        		colList.add(sortCList.get(i).getAlias());
		        	}
	        	}
	        	
	        	if(rowList.size()>0&&colList.size()>0){
	        		List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId, sheetId, rowList, colList);
	        		mcellDao.delMRowColCellList(excelId, sheetId, rowList, colList);//删除关系表
	        		Map<String,String> relationMap = new HashMap<String,String>();
	        		for(MRowColCell mrcc:relationList){
	        			relationMap.put(mrcc.getRow()+"_"+mrcc.getCol(), mrcc.getCellId());
	        		}
	        		int count=0;//找到合并区域的第一个单元格
	        		List<String> cellIdList = new ArrayList<String>();//存储需要删除的MCell的id
	        		String id = rowList.get(0)+"_"+colList.get(0);//合并单元格的id
	        		relationList.clear();//清空关系表，用于存储新创建的对象
	        		MCell mc = null;//合并单元格的MCell对象
	        		for(String row:rowList){
	        			for(String col:colList){
	        				String key = row+"_"+col;
	        				String cellId = relationMap.get(key);
	        				if(null!=cellId){
	        					if(count == 0){
	        						count++;
		        					mc = mcellDao.getMCell(excelId, sheetId, cellId);
		        				    mc.setId(id);
		        				    mc.setRowspan(rowList.size());
		        				    mc.setColspan(colList.size());
		        				    cellIdList.add(cellId);
	        					}else{
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
	        		if(count==0){
	        			mc = new MCell();
	        			mc.setId(id);
	        			mc.setSheetId(sheetId);
	        			mc.setRowspan(rowList.size());
	        			mc.setColspan(colList.size());
	        		}
	        		//删除合并区域老的MCell
	        		mcellDao.delMCell(excelId, sheetId, cellIdList);
	        		//保存新的关系表
	        		baseDao.insert(excelId, relationList);
	        		baseDao.insert(excelId, mc);
	        		//更新步骤
	        		msheetDao.updateStep(excelId, sheetId, step);
	        	}
	        	
	       }
	}
	
	@Override
	public void splitCell(Cell cell, int step, String excelId) {
		String sheetId = excelId+0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String,Integer> rowMap = new HashMap<String,Integer>();
		Map<String,Integer> colMap = new HashMap<String,Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for(int i = 0;i<sortRList.size();i++){
			rowMap.put(sortRList.get(i).getAlias(),i);
		}
		for(int i = 0;i<sortCList.size();i++){
			colMap.put(sortCList.get(i).getAlias(), i);
		}
		
		
		
		List<Coordinate> coordinates = cell.getCoordinate();
		   List<String> colList = new ArrayList<String>();
	       List<String> rowList = new ArrayList<String>();
	       
	       for(Coordinate cd:coordinates){
	        	int startRow = cd.getStartRow();
	        	int endRow = cd.getEndRow();
	        	int startCol = cd.getStartCol();
	        	int endCol = cd.getEndCol();
	        	if((endCol> -1)&&(endRow > -1)){
		        	for(int i = startRow;i<endRow+1;i++){
		        		rowList.add(sortRList.get(i).getAlias());
		        	}
		        	for(int i = startCol;i<endCol+1;i++){
		        		colList.add(sortCList.get(i).getAlias());
		        	}
	        }
	        	
	        if(rowList.size()>0&&colList.size()>0){
	        		List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId, sheetId, rowList, colList);
	        		List<String> cellIdList = new ArrayList<String>();
	        		for(MRowColCell mrcc:relationList){
	        			cellIdList.add(mrcc.getCellId());
	        		}
	        		List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId, cellIdList);
	        		List<MCell> mcList = new ArrayList<MCell>();//用于存储新生成的MCell对象
	        		cellIdList.clear();//清空，用于存贮是合并单元格的id
	        		relationList.clear();//清空，用于存储新的关系表
                    for(MCell mc:mcellList){
                    	if(mc.getColspan()>1||mc.getRowspan()>1){
                    		
                    		cellIdList.add(mc.getId());
                    		String[] ids = mc.getId().split("_");
                    		int rowIndex = rowMap.get(ids[0]);
            				int colIndex = colMap.get(ids[1]);
                    		for(int i=0;i<mc.getRowspan();i++){
                    			for(int j=0;j<mc.getColspan();j++){
                    				String row = sortRList.get(rowIndex+i).getAlias();
                    				String col = sortCList.get(colIndex+j).getAlias();
                    				MCell mcell = new MCell(mc);
                    				String id = row+"_"+col;
                    				mcell.setId(id);
                					mcell.setColspan(1);
                					mcell.setRowspan(1);
                    				if(!((i==0)&&(j==0))){
                    					mcell.getContent().setDisplayTexts(null);
                    					mcell.getContent().setTexts(null);
                    				}
                    				mcList.add(mcell);
                    				//存关系表
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
                    //删除老的合并单元格关系表
                    mcellDao.delMRowColCellList(excelId, sheetId, cellIdList);
                    //插入拆分之后的关系表
                    baseDao.insert(excelId, relationList);
                    //删除合并单元格
                    mcellDao.delMCell(excelId, sheetId, cellIdList);
                    //存入拆分之后的单元格
                    baseDao.insert(excelId, mcList);
                    //更新步骤
	        		msheetDao.updateStep(excelId, sheetId, step);
	         }
	        	
	       }
		
	}
	
	
	
	private void updateContent(List<Coordinate> coordinates,String property,Object  content,int step, String excelId,String sheetId) throws NoSuchFieldException, SecurityException, IllegalArgumentException, IllegalAccessException {
		
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		Map<String,Integer> rMap = new HashMap<String,Integer>();
		Map<String,Integer> cMap = new HashMap<String,Integer>();
		mrowColDao.getRowList(sortRList, excelId, sheetId);
		for(int i=0;i<sortRList.size();i++){
			RowCol rc = sortRList.get(i);
			rMap.put(rc.getAlias(), i);
		}
		mrowColDao.getColList(sortCList, excelId, sheetId);
		for(int i=0;i<sortCList.size();i++){
			RowCol rc = sortCList.get(i);
			cMap.put(rc.getAlias(), i);
		}
		
        List<String> colList = new ArrayList<String>();
        List<String> rowList = new ArrayList<String>();
        List<String> wholeRowList = new ArrayList<String>();
        List<String> wholeColList = new ArrayList<String>();
        for(Coordinate cd:coordinates){
        	int startRow = cd.getStartRow();
        	int endRow = cd.getEndRow();
        	int startCol = cd.getStartCol();
        	int endCol = cd.getEndCol();
        	if((endCol> -1)&&(endRow > -1)){
	        	for(int i = startRow;i<endRow+1;i++){
	        		rowList.add(sortRList.get(i).getAlias());
	        	}
	        	for(int i = startCol;i<endCol+1;i++){
	        		colList.add(sortCList.get(i).getAlias());
	        	}
        	}
        	
        	if((endCol == -1)&&(endRow > -1)){
	        	for(int i = startRow;i<endRow+1;i++){
	        		wholeRowList.add(sortRList.get(i).getAlias());
	        	}
        	}
        	
        	if((endRow == -1)&&(endCol > -1)){
        		for(int i = startCol;i<endCol+1;i++){
        			wholeColList.add(sortCList.get(i).getAlias());
	        	}
        	}
        }
        
        List<Object> tempList = new ArrayList<Object>();//用于存储新new的关系对象及MCell对象
        List<String> cellIdList = new ArrayList<String>();//用于存储需要修改单元格的id
        if(rowList.size()>0&&colList.size()>0){
        	List<MRowColCell>  list1 = mcellDao.getMRowColCellList(excelId, sheetId, rowList, colList);
        	Map<String,MRowColCell> relationMap = new HashMap<String,MRowColCell>();
        	for(MRowColCell mrcc:list1){
        		String row = mrcc.getRow();
        		String col = mrcc.getCol();
        		relationMap.put(row+"_"+col, mrcc);
        	}
        	 for(Coordinate c:coordinates){
             	int startRow = c.getStartRow();
             	int endRow = c.getEndRow();
             	int startCol = c.getStartCol();
             	int endCol = c.getEndCol();
                if(endRow>-1&&endCol>-1){
                	for(int i = startRow;i<endRow+1;i++){
                		String row = sortRList.get(i).getAlias();
                		for(int j = startCol;j<endCol+1;j++){
                		 String col = sortCList.get(j).getAlias();
                		 MRowColCell mr = relationMap.get(row+"_"+col);
                		 if(null == mr){
                			 mr = new MRowColCell();
                			 mr.setCellId(row+"_"+col);
                			 mr.setRow(row);
                			 mr.setCol(col);
                			 mr.setSheetId(sheetId);
                			 tempList.add(mr);
                			 MCell mc = new MCell(row+"_"+col,sheetId);
                			 Field f = mc.getContent().getClass().getDeclaredField(property);
                			 f.setAccessible(true);
                			 f.set(mc.getContent(), content);
                			 tempList.add(mc);
                		 }else{
                			 cellIdList.add(mr.getCellId());
                		 }
                		 
                	  }
                  }
                }
             }
        }
        
        if(wholeRowList.size()>0){
        	
        	List<MRowColCell> list2 = mcellDao.getMRowColCellList(excelId, sheetId, wholeRowList, "row");
        	Map<String,String> relationMap = new HashMap<String,String>();
        	List<String> idList = new ArrayList<String>();
        	for(MRowColCell mrcc:list2){
        		String row = mrcc.getRow();
        		String col = mrcc.getCol();
        		relationMap.put(row+"_"+col,mrcc.getCellId());
        	}
        	//获取所有的列样式对象
        	List<MCol> mcolList = mcolDao.getMColList(excelId, sheetId);
        	//用存贮选中的行
        	Map<String,String> wholeRowMap = new HashMap<String,String>();
        	for(String row:wholeRowList){
        		wholeRowMap.put(row, row);
        		for(MCol mc:mcolList){
        			String key = row+"_"+mc.getAlias();
        			String cellId = relationMap.get(key);
        			if(null == cellId){
        				if(null!=mc.getProps().getContent().getFamily()){
        					 MRowColCell mr = new MRowColCell();
                			 mr.setCellId(key);
                			 mr.setRow(row);
                			 mr.setCol(mc.getAlias());
                			 mr.setSheetId(sheetId);
                			 tempList.add(mr);//关系表
                			 MCell mcell = new MCell(key,sheetId);
                			 Field f = mcell.getContent().getClass().getDeclaredField(property);
                			 f.setAccessible(true);
                			 f.set(mcell.getContent(), content);
                			 tempList.add(mcell);//单元格对象
        				}
        			}else{
        					idList.add(cellId);
        			}
        		}
        	}
        	//找出整列操作中部分条件的合并单元格
        	List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId, idList);
        	List<ConverCell> oneCellList = getOneCellList(mcellList,rMap,cMap,sortRList,sortCList);
        	for(ConverCell oc:oneCellList){
        		List<String> rows = oc.getOccupy().getRow();
        		for(String row:rows){
        			String wholeRow = wholeRowMap.get(row);
        			if(null == wholeRow){
        			    idList.remove(oc.getId());	//剔除不符合条件的合并单元格
        			}
        		}
        	}
        	cellIdList.addAll(idList);
           //更新列对象
        	mrowDao.updateContent(property, content, wholeRowList, excelId, sheetId);
        }
        
        if(wholeColList.size()>0){
        	
        	List<MRowColCell> list3 = mcellDao.getMRowColCellList(excelId, sheetId, wholeColList, "col");
        	Map<String,String> relationMap = new HashMap<String,String>();
        	List<String> idList = new ArrayList<String>();
        	for(MRowColCell mrcc:list3){
        		String row = mrcc.getRow();
        		String col = mrcc.getCol();
        		relationMap.put(row+"_"+col,mrcc.getCellId());
        	}
        	//获取所有的行样式对象
        	List<MRow> mrowList = mrowDao.getMRowList(excelId, sheetId);
        	//存储选中列
        	Map<String,String> wholeColMap = new HashMap<String,String>();
        	for(String col:wholeColList){
        		wholeColMap.put(col, col);
        		for(MRow mr:mrowList){
        			String key = mr.getAlias()+"_"+col;
        			String cellId = relationMap.get(key);
        			if(null == cellId){
        				if(null!=mr.getProps().getContent().getFamily()){
        					 MRowColCell mrcc = new MRowColCell();
        					 mrcc.setCellId(key);
        					 mrcc.setRow(mr.getAlias());
        					 mrcc.setCol(col);
        					 mrcc.setSheetId(sheetId);
                			 tempList.add(mrcc);//关系表
                			 MCell mcell = new MCell(key,sheetId);
                			 Field f = mcell.getContent().getClass().getDeclaredField(property);
                			 f.setAccessible(true);
                			 f.set(mcell.getContent(), content);
                			 
                			 tempList.add(mcell);//单元格对象
        				}
        			}else{
        					idList.add(cellId);
        			}
        		}
        	}
        	//找出整列操作中部分条件的合并单元格
        	List<MCell> mcellList = mcellDao.getMCellList(excelId, sheetId, idList);
        	List<ConverCell> oneCellList = getOneCellList(mcellList,rMap,cMap,sortRList,sortCList);
        	for(ConverCell oc:oneCellList){
        		List<String> cols = oc.getOccupy().getCol();
        		for(String col:cols){
        			String wholeRow = wholeColMap.get(col);
        			if(null == wholeRow){
        			    idList.remove(oc.getId());	//剔除不符合条件的合并单元格
        			}
        		}
        	}
        	cellIdList.addAll(idList);
        	
           //更新列对象
        	mcolDao.updateContent(property, content, wholeColList, excelId, sheetId);
        }
        //更新MCell对象属性
        mcellDao.updateContent(property, content, cellIdList, excelId, sheetId);
        baseDao.insert(excelId, tempList);//存储新创建的关系表及MCell对象
        msheetDao.updateStep(excelId, sheetId, step);
		
	}
	
	private List<ConverCell> getOneCellList(List<MCell> cellList,Map<String,Integer> rMap,Map<String,Integer> cMap,
			List<RowCol> sortRList,List<RowCol> sortCList){
		List<ConverCell> list = new ArrayList<>();
		for(MCell mc:cellList){
			
			if(mc.getRowspan()==1&&mc.getColspan()==1){
				
			}else{
				ConverCell  ce = new ConverCell();
				ce.setId(mc.getId());
				String[] ids = mc.getId().split("_");
				Occupy oc = ce.getOccupy();
				for(int i=0;i<mc.getRowspan();i++){
					int index = rMap.get(ids[0]);
					oc.getRow().add(sortRList.get(index).getAlias());
					index++;
				}
				for(int i=0;i<mc.getColspan();i++){
					int index = cMap.get(ids[1]);
					oc.getCol().add(sortCList.get(index).getAlias());
					index++;
				}
			  list.add(ce);
			}
			
			
		}
		return list;
	}

	
	
	
	/*public void OperationCellByCord(ArrayList<Object> al){
		//查找关系表
	    List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId,sheetId,alias,"col");
	    List<String> cellIdList = new ArrayList<String>();
	    //Map<String,String> relationMap = new HashMap<String,String>();
	    for(MRowColCell mrcc:relationList){
	    	
	    	//relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
	    	cellIdList.add(mrcc.getCellId());
	    }
        List<MCell> cellList = mcellDao.getMCellList(excelId,sheetId, cellIdList);
        //删除关系表
        mcellDao.delMRowColCell(excelId,sheetId,"col", alias);
        cellIdList.clear();//存需要删除的MExcelCell的Id
		cellList.clear();//存需要插入的MExcelCell对象
		
		for(MCell mc:cellList){
			if(mc.getColspan()==1){
				cellIdList.add(mc.getId());
			}else{
				String[] ids = mc.getId().split("_");
				if(ids[1].equals(alias)){
					//删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					String id = ids[0]+"_"+backAlias;
					//修改合并单元格其他关系表的cellId
					mcellDao.updateMRowColCell(excelId,sheetId, mec.getId(), id);
					mec.setId(id);
					mec.setColspan(mc.getColspan()-1);
					
					cellList.add(mec);//插入新MCell
				}else{
					MCell mec = mc;
					mec.setColspan(mc.getColspan()-1);
					baseDao.update(excelId, mec);
				}
			}
		}
		
		mcellDao.delMCell(excelId,sheetId, cellIdList);
		if(cellList.size()>0){
		 baseDao.insert(excelId, cellList);
		
		}
		
	}*/
	
}
