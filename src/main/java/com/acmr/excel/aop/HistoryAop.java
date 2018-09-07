package com.acmr.excel.aop;

import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.CopyOnWriteArrayList;

import javax.annotation.Resource;
import javax.servlet.http.HttpServletRequest;

import org.aspectj.lang.JoinPoint;
import org.aspectj.lang.annotation.Aspect;
import org.aspectj.lang.annotation.Before;
import org.aspectj.lang.annotation.Pointcut;
import org.springframework.stereotype.Component;

import com.acmr.excel.dao.MCellDao;
import com.acmr.excel.dao.MColDao;
import com.acmr.excel.dao.MRowColCellDao;
import com.acmr.excel.dao.MRowColDao;
import com.acmr.excel.dao.MRowDao;
import com.acmr.excel.dao.base.BaseDao;
import com.acmr.excel.model.Constant;
import com.acmr.excel.model.history.History;
import com.acmr.excel.model.history.HistoryCache;
import com.acmr.excel.model.history.MCellBefore;
import com.acmr.excel.model.history.MColBefore;
import com.acmr.excel.model.history.MRowBefore;
import com.acmr.excel.model.history.MRowColCellBefore;
import com.acmr.excel.model.history.Param;
import com.acmr.excel.model.history.Record;
import com.acmr.excel.model.mongo.MCell;
import com.acmr.excel.model.mongo.MCol;
import com.acmr.excel.model.mongo.MRow;
import com.acmr.excel.model.mongo.MRowColCell;
import com.acmr.excel.model.mongo.MRowColList;
import com.acmr.excel.model.mongo.MSheet;
import com.acmr.redis.Redis;



public class HistoryAop {

	private  CopyOnWriteArrayList<History> list;

	/* 当前操作位置 */
	private  int index;
	
	private HistoryCache cache;

	private History history;

	private Record record;

	private String methodName;

	// 存储修改过属性行样式id
	private List<String> rowIdList;
	// 存放修改前mrow对象
	private List<MRow> mrowList;

	private List<String> colIdList;
	private List<MCol> mcolList;

	private List<String> cellIdList;
	private List<MCell> mcellList;

	private List<String> mrowColCellIdList;
	private List<MRowColCell> mrowColCellList;

	private List<MRowColList> rowList;
	private List<MRowColList> colList;

	private List<Object> delList;

	private String excelId;
	private String sheetId;

	@Resource
	private MRowColDao mrowColDao;
	@Resource
	private BaseDao baseDao;
	@Resource
	private MCellDao mcellDao;
	@Resource
	private MColDao mcolDao;
	@Resource
	private MRowDao mrowDao;
	@Resource
	private MRowColCellDao mrowColCellDao;
	@Resource
	private Redis redis;

	@Pointcut("execution(* com.acmr.excel.controller.*.*(..))")
	public void pointCut() {
	}

	@Pointcut("execution(* com.acmr.excel.dao..*.*(..))")
	public void pointCut1() {
	}
	

	@SuppressWarnings("unchecked")
	@Before("pointCut()")
	public void beforeAction(JoinPoint joinPoint) {
		System.out.println("进入切点");
		
		methodName = joinPoint.getSignature().getName();
		System.out.println(joinPoint.getStaticPart().toShortString());
		System.out.println(joinPoint.getSourceLocation().toString());
		

		Object[] arg = joinPoint.getArgs();
		HttpServletRequest req = (HttpServletRequest) arg[0];
		excelId = req.getHeader("X-Book-Id");
		if(null==excelId){
			return;
		}
		sheetId = excelId + 0;
		cache = (HistoryCache) redis.get(excelId);
		if(null == cache){
			list = new CopyOnWriteArrayList<History>();
			cache = new HistoryCache();
		}else{
		   	list = cache.getList();
			
			//当中断前进后退操作时，删除这部之后所有的记录
			if(!("undo".equals(methodName)||"redo".equals(methodName))){
				
				/*index = cache.getIndex();
				for(int i=index;i<list.size();i++){
					list.remove(i);
				}*/
				cache.setList(list);
				redis.set(excelId, cache);
			}
			
	   }
		
		if (null != Constant.unRecordAction.get(methodName)) {
			return;
		}

		
		

		rowIdList = new ArrayList<String>();
		mrowList = new ArrayList<MRow>();

		colIdList = new ArrayList<String>();
		mcolList = new ArrayList<MCol>();

		cellIdList = new ArrayList<String>();
		mcellList = new ArrayList<MCell>();

		mrowColCellIdList = new ArrayList<String>();
		mrowColCellList = new ArrayList<MRowColCell>();

		delList = new ArrayList<Object>();

		rowList = new ArrayList<MRowColList>();
		colList = new ArrayList<MRowColList>();

		history = new History();
		record = history.getRecord();
		history.setName(methodName);
		history.setExcelId(excelId);
		MRowBefore mrow = new MRowBefore(rowIdList, mrowList);
		MColBefore mcol = new MColBefore(colIdList, mcolList);
		MCellBefore mcell = new MCellBefore(cellIdList, mcellList);
		MRowColCellBefore mrowColCell = new MRowColCellBefore(mrowColCellIdList,
				mrowColCellList);
		com.acmr.excel.model.history.Before before = history.getBefore();
		before.setMrow(mrow);
		before.setMcol(mcol);
		before.setMcell(mcell);
		before.setMrowColCell(mrowColCell);
		before.setColList(colList);
		before.setRowList(rowList);
		before.setDelList(delList);
        
		list.add(history);
		index=list.size();
		cache.setList(list);
		cache.setIndex(index);
		
		
	}

	@Before("pointCut1()")
	public void beforeDao(JoinPoint joinPoint) {
		if(null==excelId){
			return;
		}
		
		if ((null == methodName)||(null != Constant.unRecordAction.get(methodName))) {
			return;
		}

		String target = joinPoint.getTarget().getClass().getSimpleName();
		Object[] arg = joinPoint.getArgs();
		String name = joinPoint.getSignature().getName();
		String shortName = joinPoint.getStaticPart().toShortString();
		if (shortName.indexOf("MSheetDao") > -1 || name.indexOf("get") > -1
				|| (shortName.indexOf("BaseDao") > -1
						&& arg[1].getClass() == MSheet.class)) {
			return;
		}

		Param param = new Param(target, name, arg);
		record.getParam().add(param);

		recordBeford(shortName, name, arg);
		
		redis.set(excelId, cache);//存入redis

	}

	private void recordBeford(String shortName, String methodName,
			Object[] arg) {
		if (shortName.indexOf("MRowDao") > -1) {
			if ("updateContent".equals(methodName)) {
				rowIdList.addAll((ArrayList) arg[4]);
				List<MRow> list = mrowDao.getMRowList(excelId, excelId + 0,
						(ArrayList) arg[4]);
				mrowList.addAll(list);
			} else {
				if (arg[2].getClass() == String.class) {
					rowIdList.add((String) arg[2]);
					MRow mrow = mrowDao.getMRow(excelId, excelId + 0,
							(String) arg[2]);
					mrowList.add(mrow);
				} else {
					rowIdList.addAll((ArrayList) arg[2]);
					List<MRow> list = mrowDao.getMRowList(excelId, excelId + 0,
							(ArrayList) arg[4]);
					mrowList.addAll(list);
				}
			}
		} else if (shortName.indexOf("MColDao") > -1) {
			if ("updateContent".equals(methodName)) {
				colIdList.addAll((ArrayList) arg[4]);
				List<MCol> list = mcolDao.getMColList(excelId, excelId + 0,
						(ArrayList) arg[4]);
				mcolList.addAll(list);
			} else {
				if (arg[2].getClass() == String.class) {
					colIdList.add((String) arg[2]);
					MCol mcol = mcolDao.getMCol(excelId, excelId + 0,
							(String) arg[2]);
				} else {
					colIdList.addAll((ArrayList) arg[2]);
					List<MCol> list = mcolDao.getMColList(excelId, excelId + 0,
							(ArrayList) arg[4]);
					mcolList.addAll(list);
				}
			}
		} else if (shortName.indexOf("MRowColDao") > -1) {
			if ("updateRowCol".equals(methodName)
					|| "updateRowColLength".equals(methodName)) {
				String id = (String) arg[2];
				if ("rList".equals(id) && rowList.size() < 1) {
					MRowColList rows = mrowColDao.getMRowColList("rList",
							excelId, sheetId);
					rowList.add(rows);
				} else if ("cList".equals(id) && colList.size() < 1) {
					MRowColList cols = mrowColDao.getMRowColList("cList",
							excelId, sheetId);
					colList.add(cols);
				}
			} else if ("insertRowCol".equals(methodName)
					|| "delRowCol".equals(methodName)) {
				String id = (String) arg[3];
				if ("rList".equals(id) && rowList.size() < 1) {
					MRowColList rows = mrowColDao.getMRowColList("rList",
							excelId, sheetId);
					rowList.add(rows);
				} else if ("cList".equals(id) && colList.size() < 1) {
					MRowColList cols = mrowColDao.getMRowColList("cList",
							excelId, sheetId);
					colList.add(cols);
				}
			}
		} else if (shortName.indexOf("MCellDao") > -1) {
			if ("updateContent".equals(methodName)) {
				cellIdList.addAll((ArrayList) arg[4]);
				List<MCell> list = mcellDao.getMCellList(excelId, excelId + 0,
						(ArrayList) arg[4]);
				mcellList.addAll(list);
			} else {
				if (arg[2].getClass() == String.class) {
					cellIdList.add((String) arg[2]);
					MCell mcell = mcellDao.getMCell(excelId, excelId + 0,
							(String) arg[2]);
					mcellList.add(mcell);
				} else {
					cellIdList.addAll((ArrayList) arg[2]);
					List<MCell> list = mcellDao.getMCellList(excelId,
							excelId + 0, (ArrayList) arg[2]);
					mcellList.addAll(list);
				}
			}
		} else if (shortName.indexOf("MRowColCellDao") > -1) {
			if ("delMRowColCell".equals(methodName)) {
				List<MRowColCell> list = mrowColCellDao.getMRowColCellList(
						excelId, excelId + 0, (String) arg[3], (String) arg[2]);
				mrowColCellList.addAll(list);
			} else if (("delMRowColCellList".equals(methodName))
					&& (arg.length == 4)) {
				List<MRowColCell> list = mrowColCellDao.getMRowColCellList(
						excelId, excelId + 0, (ArrayList) arg[2],
						(ArrayList) arg[3]);
				mrowColCellList.addAll(list);
			} else if ("delMRowColCellList1".equals(methodName)) {
				List<MRowColCell> list = mrowColCellDao.getMRowColCellList(
						excelId, excelId + 0, (ArrayList) arg[2]);
				mrowColCellList.addAll(list);
			} else if ("updateMRowColCell".equals(methodName)) {
				mrowColCellIdList.add((String) arg[2]);
				List<MRowColCell> list = mrowColCellDao.getMRowColCellList(
						excelId, excelId + 0, (String) arg[2]);
				mrowColCellList.addAll(list);
			}
		} else if (shortName.indexOf("BaseDao") > -1) {
			if ("insert".equals(methodName)||"insertList".equals(methodName)) {
				if (arg[1].getClass() == ArrayList.class) {
					delList.addAll((ArrayList) arg[1]);
				} else {
					delList.add(arg[1]);
				}
			} else if ("update".equals(methodName)) {
				if (arg[1].getClass() == MCell.class) {
					String id = ((MCell) arg[1]).getId();
					cellIdList.add(id);
					MCell mc = mcellDao.getMCell(excelId, excelId + 0, id);
					mcellList.add(mc);
				}

			}
		}
	}

}
