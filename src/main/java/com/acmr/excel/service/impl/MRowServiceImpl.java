package com.acmr.excel.service.impl;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import javax.annotation.Resource;

import org.springframework.stereotype.Service;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MExcelRowDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.RowCol;
import com.acmr.excel.model.RowColCell;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowOperate;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MExcelCell;
import com.acmr.excel.model.mongo.MExcelRow;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.excel.service.MRowService;

import acmr.excel.pojo.ExcelCell;

@Service("mrowService")
public class MRowServiceImpl implements MRowService {

	@Resource
	private BaseDao baseDao;
	
	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MExcelRowDao mexcelRowDao;
	@Resource
	private MSheetDao msheetDao;
	@Resource
	private MRowDao mrowDao;
	
	public void insertRow(RowOperate rowOperate,String excelId,Integer step){
		String sheetId = excelId+0;
		List<RowCol> sortRcList = new ArrayList<RowCol>();
		List<RowCol> sortClList = new ArrayList<RowCol>();
		mrowColDao.getColList(sortClList, excelId,sheetId);
		mrowColDao.getRowList(sortRcList, excelId,sheetId);
		if(rowOperate.getRow()>sortRcList.size()-1){
			return;
		}
		Map<String,Integer> colMap = new HashMap<String,Integer>();
		for(int i = 0;i<sortClList.size();i++){
			RowCol rc = sortClList.get(i);
			colMap.put(rc.getAlias(), i);
		}
		
		MSheet msheet = msheetDao.getMSheet(excelId,sheetId);
		String alias = msheet.getMaxrow()+1+"";
		
		RowCol rowCol = new RowCol();
		rowCol.setAlias(alias);
		rowCol.setLength(19);
		if(rowOperate.getRow() == 0){
			rowCol.setPreAlias(null);
		}else{
			rowCol.setPreAlias(sortRcList.get(rowOperate.getRow()-1).getAlias());
		}
		
		mrowColDao.insertRowCol(excelId,sheetId, rowCol, "rList");
		//修改这行的前行别名
		String nextAlias = sortRcList.get(rowOperate.getRow()).getAlias();
		mrowColDao.updateRowCol(excelId,sheetId, "rList", nextAlias, alias);
		//存一个行样式
		MRow mr = new MRow();
		mr.setAlias(alias);
		mr.setHeight(19);
		mr.setHidden(false);
		mr.setSheetId(sheetId);
		
		baseDao.insert(excelId, mr);
		
		msheetDao.updateMaxRow(msheet.getMaxrow()+1, excelId,sheetId);
		
		String  row  = sortRcList.get(rowOperate.getRow()).getAlias();
		List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId,sheetId, row,"row");
		List<String> cellIdList = new ArrayList<>();
		//Map<String,String> relationMap = new HashMap<String,String>();
		for(MRowColCell mrcc:relationList){
			cellIdList.add(mrcc.getCellId());
		//	relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
		}
		List<MCell> mcellList = mcellDao.getMCellList(excelId,sheetId, cellIdList);
		List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		for(MCell mc : mcellList){
			String[] ids = mc.getId().split("_");
			if((mc.getRowspan()>1)&&(!ids[0].equals(row))){
				mc.setColspan(mc.getColspan()+1);
			}
			tempList.add(mc);
			for(int i=0;i<mc.getColspan();i++){
				MRowColCell mrcc = new MRowColCell();
				String col = ids[1];
				Integer index = colMap.get(col);
				row = sortRcList.get(index).getAlias();
				index++;
				mrcc.setCellId(mc.getId());
				mrcc.setRow(alias);
				mrcc.setCol(col);
				mrcc.setSheetId(sheetId);
				tempList.add(mrcc);
			}
			
		}
		
		if(tempList.size()>0){
			baseDao.update(excelId, tempList);
		}
		
		msheetDao.updateStep(excelId,sheetId, step);
		
		
		/*Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
		for(MExcelCell mec:mcellList){
			cellMap.put(mec.getId(), mec);
		}
		String rowId = sortRcList.get(rowOperate.getRow()).getAlias();//行别名
		List<MExcelCell> sortMCellList = new ArrayList<MExcelCell>();
		
		for(RowCol rc: sortClList){
			String cellId = relationMap.get(rowId+"_"+rc.getAlias());
			sortMCellList.add(cellMap.get(cellId));
		}
		
		List<Object> tempList = new ArrayList<Object>();//存关系表和MExcelCell对象
		for(int i=0;i<sortClList.size();i++){
			MExcelCell mec = sortMCellList.get(i);
			if(null==mec){
				continue;
			}
			RowCol rc = sortClList.get(i);
			if(mec.getRowspan()==1){
				//new新增一行的MExcelCell对象
				MExcelCell me = new MExcelCell();
				me.setId(alias+"_"+rc.getAlias());
				me.setRowId(alias);
				me.setColId(rc.getAlias());
				me.setRowspan(1);
				me.setColspan(1);
				ExcelCell ec = new ExcelCell();
				me.setExcelCell(ec);
				tempList.add(me);
				RowColCell rcc = new RowColCell();
				rcc.setCellId(alias+"_"+rc.getAlias());
				rcc.setCol(Integer.parseInt(rc.getAlias()));
				rcc.setRow(Integer.parseInt(alias));
				tempList.add(rcc);
			}else if(mec.getRowspan()>1){
				if(mec.getRowId().equals(rowId) ){
					//new新增一行的MExcelCell对象,是以选定行作为开始行的合并单元格
					MExcelCell me = new MExcelCell();
					me.setId(alias+"_"+rc.getAlias());
					me.setRowId(alias);
					me.setColId(rc.getAlias());
					me.setRowspan(1);
					me.setColspan(1);
					ExcelCell ec = new ExcelCell();
					me.setExcelCell(ec);
					tempList.add(me);
					//关系对象
					RowColCell rcc = new RowColCell();
					rcc.setCellId(alias+"_"+rc.getAlias());
					rcc.setCol(Integer.parseInt(rc.getAlias()));
					rcc.setRow(Integer.parseInt(alias));
					tempList.add(rcc);
				}else{
					if(mec.getColId().equals(rc.getAlias())){
						mec.setRowspan(mec.getRowspan()+1);
						tempList.add(mec);
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(rc.getAlias()));
						rcc.setRow(Integer.parseInt(alias));
						tempList.add(rcc);
					}else{
						//关系对象
						RowColCell rcc = new RowColCell();
						rcc.setCellId(mec.getId());
						rcc.setCol(Integer.parseInt(rc.getAlias()));
						rcc.setRow(Integer.parseInt(alias));
						tempList.add(rcc);
					}
				}
			}
		}
		*/
		
	}

	@Override
	public void delRow(RowOperate rowOperate, String excelId,Integer step) {
		String sheetId = excelId+0;
		List<RowCol> sortRList = new ArrayList<RowCol>();
		List<RowCol> sortCList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId,"");
		mrowColDao.getColList(sortCList, excelId,"");
		if(rowOperate.getRow()>sortRList.size()-1){
			return;
		}
		//行别名
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		//删除行
		mrowColDao.delRowCol(excelId,sheetId, alias, "rList");
		//修改这行后面一行的前一行别名
		int index = rowOperate.getRow();
		String frontAlias;
		if(rowOperate.getRow() == 0){
			frontAlias = null;//前一行别名
		}else{
			frontAlias = sortRList.get(index-1).getAlias();//前一行别名
		}
		 
		String backAlias = sortRList.get(index+1).getAlias();//下一行别名
		mrowColDao.updateRowCol(excelId,sheetId, "rList", backAlias, frontAlias);
		//删除行样式记录
		mrowDao.delMRow(excelId,sheetId, alias);
	    List<MRowColCell> relationList = mcellDao.getMRowColCellList(excelId,sheetId, alias,"row");
	    List<String> cellIdList = new ArrayList<String>();
	    //Map<String,String> relationMap = new HashMap<String,String>();
	    for(MRowColCell mrcc:relationList){
	    	
	    	//relationMap.put(rcc.getRow()+"_"+rcc.getCol(), rcc.getCellId());
	    	cellIdList.add(mrcc.getCellId());
	    }
        List<MCell> cellList = mcellDao.getMCellList(excelId,sheetId, cellIdList);
       //删除关系表
        mcellDao.delMRowColCell(excelId,sheetId,"row",alias);
      	cellIdList.clear();//存需要删除的MExcelCell的Id
      	cellList.clear();//存需要修改或增加的MExcelCell对象
      	for(MCell mc:cellList){
			if(mc.getRowspan()==1){
				cellIdList.add(mc.getId());
			}else{
				String[] ids = mc.getId().split("_");
				if(ids[0].equals(alias)){
					//删除老的MExcelCell
					cellIdList.add(mc.getId());
					MCell mec = mc;
					String id = backAlias+"_"+ids[1];
					//修改合并单元格其他关系表的cellId
					mcellDao.updateMRowColCell(excelId,sheetId, mec.getId(), id);
					mec.setId(id);
					mec.setRowspan(mc.getRowspan()-1);
					
					cellList.add(mec);//插入新MCell
				}else{
					MCell mec = mc;
					mec.setColspan(mc.getColspan()-1);
					cellList.add(mec);
				}
			}
		}
      	
      	mcellDao.delMCell(excelId,sheetId, cellIdList);
		if(cellList.size()>0){
			baseDao.update(excelId, cellList);
		}
		msheetDao.updateStep(excelId,sheetId, step);
      	
       /* List<MExcelCell> sortMcellList = new ArrayList<MExcelCell>();
        Map<String,MExcelCell> cellMap = new HashMap<String,MExcelCell>();
        for(MExcelCell ec:cellList){
        	cellMap.put(ec.getId(), ec);
        }
		for(RowCol rc:sortCList){
			String cellId = relationMap.get(alias+"_"+rc.getAlias());
			sortMcellList.add(cellMap.get(cellId));
		}
		//删除关系表
		mcellDao.delRowColCell(excelId,"row",Integer.parseInt(alias));
		cellIdList.clear();//存需要删除的MExcelCell的Id
		cellList.clear();//存需要修改或增加的MExcelCell对象
		for(int i=0;i<sortCList.size();i++){
			String col = sortCList.get(i).getAlias();
			MExcelCell ec = sortMcellList.get(i);
			if(null == ec){
				continue;
			}
			if(ec.getRowspan() == 1){
				cellIdList.add(ec.getId());
			}else{
				if(ec.getRowId().equals(alias)){
					if(ec.getColId().equals(col)){
						//删除老的MExcelCell
						cellIdList.add(ec.getId());
						MExcelCell mec = ec;
						mec.setRowId(backAlias);
						String id = backAlias+"_"+col;
						//修改合并单元格其他关系表的cellId
						mcellDao.updateRowColCell(excelId, mec.getId(), id);
						mec.setId(backAlias+"_"+col);
						mec.setRowspan(ec.getRowspan()-1);
						
						ExcelCell cell = ec.getExcelCell();
						cell.setRowspan(cell.getRowspan()-1);
						mec.setExcelCell(cell);
						cellList.add(mec);//插入新MExcelCell
								
					}
				}else{
					if(ec.getColId().equals(col)){
						MExcelCell mec = ec;
						mec.setRowspan(mec.getRowspan()-1);
						cellList.add(mec);
					}
				} 
			}
			
		}*/
		
	}

	@Override
	public void hideRow(RowOperate rowOperate, String excelId,Integer step) {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId,"");
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		mexcelRowDao.updateRowHidden(excelId, alias, true);
		mrowColDao.updateRowColLength(excelId, "rList", alias, 0);
		//mcellDao.updateHidden("rowId", alias, true, excelId);
		msheetDao.updateStep(excelId,"", step);
		
	}

	@Override
	public void showRow(RowOperate rowOperate, String excelId,Integer step) {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId,"");
		String alias = sortRList.get(rowOperate.getRow()).getAlias();
		mexcelRowDao.updateRowHidden(excelId, alias, false);
		
		MExcelRow mrow = mexcelRowDao.getMExcelRow(excelId, alias);
		mrowColDao.updateRowColLength(excelId, "rList", alias, mrow.getExcelRow().getHeight());
		
		//mcellDao.updateHidden("rowId", alias, false, excelId);
		msheetDao.updateStep(excelId,"", step);
		
	}

	@Override
	public void updateRowHeight(RowHeight rowHeight, String excelId, Integer step) {
		List<RowCol> sortRList = new ArrayList<RowCol>();
		mrowColDao.getRowList(sortRList, excelId,"");
		String alias = sortRList.get(rowHeight.getRow()).getAlias();
		Integer height = rowHeight.getOffset();
		mexcelRowDao.updateRowHeight(excelId, alias,height);
		mrowColDao.updateRowColLength(excelId, "rList", alias, height);
		msheetDao.updateStep(excelId,"", step);
		
	}

}
