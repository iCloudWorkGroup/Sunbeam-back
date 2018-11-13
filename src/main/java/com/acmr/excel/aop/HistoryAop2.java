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
import org.springframework.context.annotation.Scope;
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

//@Aspect
@Component
public class HistoryAop2 {

	private ThreadLocal<HistoryCache> thCache;

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

	@Pointcut("execution(* com.acmr.excel.service.*.*(..)) && (!within(com.acmr.excel.service.impl.MBookServiceImpl)) &&"
			+ "(!execution(* com.acmr.excel.service.impl.MSheetServiceImpl.getStep(..)))"
			+ "&&(!execution(* com.acmr.excel.service.impl.MSheetServiceImpl.updateStep(..)))")
	public void pointCut2() {
	}

	public void beforeAction(JoinPoint joinPoint) {

	}

	@Before("pointCut2()")
	public void beforeService(JoinPoint joinPoint) {
		System.out.println("进入切点");
		String methodName = joinPoint.getSignature().getName();
		System.out.println(joinPoint.getStaticPart().toShortString());
		System.out.println(methodName);

		if (methodName.indexOf("saveExcelBook") > -1
				|| methodName.indexOf("reloadExcelBook") > -1
				|| methodName.indexOf("getExcelBook") > -1
				|| methodName.indexOf("getExcels") > -1
				|| methodName.indexOf("getStep") > -1
				|| methodName.indexOf("updateStep") > -1
				|| methodName.indexOf("isAblePaste") > -1
				|| methodName.indexOf("isCut") > -1
				|| methodName.indexOf("isCopy") > -1
				|| methodName.indexOf("isCut") > -1) {
			return;
		}

		Object[] arg = joinPoint.getArgs();
		String bookId;
		try {
			bookId = (String) arg[0];
		} catch (Exception e) {
			return;
		}

		HistoryCache cache = (HistoryCache) redis.get(bookId);
		CopyOnWriteArrayList<History> list = null;
		if (null == cache) {
			cache = new HistoryCache();
			list = cache.getList();
		} else {
			list = cache.getList();

			// 当中断前进后退操作时，删除这部之后所有的记录
			if (!("undo".equals(methodName) || "redo".equals(methodName))) {

				/*
				 * index = cache.getIndex(); for(int i=index;i<list.size();i++){
				 * list.remove(i); }
				 */
				cache.setList(list);
				redis.set(bookId, cache);
			}

		}

		if (null != Constant.unRecordAction.get(methodName)) {
			return;
		}

		History history = new History();
		history.setName(methodName);
		history.setExcelId(bookId);
		list.add(history);

		int index = list.size();
		cache.setList(list);
		cache.setIndex(index);

		thCache = new ThreadLocal<HistoryCache>();
		thCache.set(cache);// 放入线程局部变量载体

		redis.set(bookId, cache);
	}

	@Before("pointCut1()")
	public void beforeDao(JoinPoint joinPoint) {

		if (null == thCache) {
			return;
		}

		String target = joinPoint.getTarget().getClass().getSimpleName();// 目标类名
		Object[] arg = joinPoint.getArgs();
		String name = joinPoint.getSignature().getName();// 方法名
		String shortName = joinPoint.getStaticPart().toShortString();// 类名
		if (shortName.indexOf("MSheetDao") > -1 || name.indexOf("get") > -1
				|| (shortName.indexOf("BaseDao") > -1
						&& arg[1].getClass() == MSheet.class)) {
			return;
		}
		HistoryCache cache = thCache.get();// 从线程局部变量中拿出对象
		
		if (null == cache) {
			return;
		}

		History history = cache.getList().get(cache.getIndex() - 1);

		Param param = new Param(target, name, arg);
		history.getRecord().getParam().add(param);

		recordBeford(shortName, name, arg, history);

		thCache.set(cache);

		redis.set(history.getExcelId(), cache);// 存入redis

	}

	private void recordBeford(String shortName, String methodName, Object[] arg,
			History history) {
		com.acmr.excel.model.history.Before before = history.getBefore();
		String excelId = history.getExcelId();
		String sheetId = excelId + 0;

		MCellBefore mcellBefore = before.getMcell();
		List<String> cellIdList = mcellBefore.getIdList();
		List<MCell> mcellList = mcellBefore.getMcellList();

		if (shortName.indexOf("MRowDao") > -1) {

			MRowBefore mrowBefore = before.getMrow();
			List<String> rowIdList = mrowBefore.getIdList();
			List<MRow> mrowList = mrowBefore.getMrowList();

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
							(ArrayList) arg[2]);
					mrowList.addAll(list);
				}
			}
		} else if (shortName.indexOf("MColDao") > -1) {

			MColBefore mcolBefore = before.getMcol();
			List<String> colIdList = mcolBefore.getIdList();
			List<MCol> mcolList = mcolBefore.getMcolList();

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
							(ArrayList) arg[2]);
					mcolList.addAll(list);
				}
			}
		} else if (shortName.indexOf("MRowColDao") > -1) {

			List<MRowColList> rowList = before.getRowList();
			List<MRowColList> colList = before.getColList();

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

			MRowColCellBefore mrowColCellBefore = before.getMrowColCell();
			List<MRowColCell> mrowColCellList = mrowColCellBefore
					.getMrowColCellList();
			List<String> mrowColCellIdList = mrowColCellBefore.getIdList();

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

			List<Object> delList = before.getDelList();

			if ("insert".equals(methodName)
					|| "insertList".equals(methodName)) {
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
