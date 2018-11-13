package com.acmr.mq.consumer.queue;

import org.apache.log4j.Logger;

import com.acmr.excel.dao.MSheetDao;
import com.acmr.excel.model.Cell;
import com.acmr.excel.model.CellContent;
import com.acmr.excel.model.CellFormate;
import com.acmr.excel.model.ColOperate;
import com.acmr.excel.model.ColWidth;
import com.acmr.excel.model.Frozen;
import com.acmr.excel.model.OperatorConstant;
import com.acmr.excel.model.OuterPaste;
import com.acmr.excel.model.RowHeight;
import com.acmr.excel.model.RowOperate;
import com.acmr.excel.model.RowOrColExpand;
import com.acmr.excel.model.copy.Copy;
import com.acmr.excel.service.MCellService;
import com.acmr.excel.service.MColService;
import com.acmr.excel.service.MRowService;
import com.acmr.excel.service.MSheetService;
import com.acmr.mq.Model;


public class WorkerThread2 implements Runnable {
	private static Logger logger = Logger.getLogger(QueueReceiver.class);
	private int step;
	private String key;
	private Model model;

	private MCellService mcellService;
	
	private MRowService mrowService;
	
	private MColService mcolService;
	
	private MSheetService msheetService;
	
	private MSheetDao msheetDao;

	public WorkerThread2(int step, String key, Model model,
			MCellService mcellService, MRowService mrowService,
			MColService mcolService, MSheetService msheetService,
			MSheetDao msheetDao) {

		this.step = step;
		this.key = key;
		this.model = model;
		this.mcellService = mcellService;
		this.mrowService = mrowService;
		this.mcolService = mcolService;
		this.msheetService = msheetService;
		this.msheetDao = msheetDao;
	}
	
	public WorkerThread2(int step, String key, Model model) {

		this.step = step;
		this.key = key;
		this.model = model;
	}
	
	public WorkerThread2() {

	}
	

	@Override
	public void run() {
		while (true) {
			int memStep = msheetService.getStep(key, key+0);
			if (memStep + 1 == step) {
				System.out.println(step + "开始执行");
				logger.info("begin excelId:" + model.getExcelId() + ";step:"
						+ step + ";reqPath:" + model.getReqPath());

				handleMessage(model);
				return;
			} else {
				processCommand(10);
				continue;
			}
		}
	}

	private void processCommand(int n) {
		try {
			Thread.sleep(n);
		} catch (InterruptedException e) {
			e.printStackTrace();
		}
	}

	@Override
	public String toString() {
		return this.step + "";
	}

	public void handleMessage(Model model) {
		int reqPath = model.getReqPath();
		String excelId = model.getExcelId();
		int step = model.getStep();
		Cell cell = null;
		switch (reqPath) {
		case OperatorConstant.textData:
			CellContent cellContent = (CellContent) model.getObject();
			mcellService.saveContent(excelId,cellContent, step);
			break;
		case OperatorConstant.clearCell:
			cell = (Cell) model.getObject();
			mcellService.clearContent(excelId,cell, step);
			break;
		case OperatorConstant.fontsize:
			cell = (Cell) model.getObject();
			mcellService.updateFontSize(excelId,cell, step);
			break;
		case OperatorConstant.fontfamily:
			cell = (Cell) model.getObject();
			mcellService.updateFontFamily(excelId,cell, step);
			break;
		case OperatorConstant.fontweight:
			cell = (Cell) model.getObject();
			mcellService.updateFontWeight(excelId,cell, step);
			break;
		case OperatorConstant.fontitalic:
			cell = (Cell) model.getObject();
			mcellService.updateFontItalic(excelId,cell, step);
			break;
		case OperatorConstant.fontcolor:
			cell = (Cell) model.getObject();
			mcellService.updateFontColor(excelId,cell, step);
			break;
		case OperatorConstant.wordWrap:
			cell = (Cell) model.getObject();
			mcellService.updateWordwrap(excelId,cell, step);
			break;

		case OperatorConstant.fillbgcolor:
			cell = (Cell) model.getObject();
			mcellService.updateBgColor(excelId,cell, step);
			break;
		case OperatorConstant.textDataformat:
			CellFormate cellFormate = (CellFormate) model.getObject();
			mcellService.updateFormat(excelId,cellFormate, step);
			break;
		case OperatorConstant.underline:
			cell = (Cell) model.getObject();
			mcellService.updateFontUnderline(excelId,cell, step);
			break;
		 case OperatorConstant.commentset:
		   cell = (Cell) model.getObject();
		   mcellService.updateComment(excelId,cell, step);
		   break;
		case OperatorConstant.merge:
			cell = (Cell) model.getObject();
			mcellService.mergeCell(excelId,cell, step);
			break;
		case OperatorConstant.mergedelete:
			cell = (Cell) model.getObject();
			mcellService.splitCell(excelId,cell, step);
			break;
		case OperatorConstant.frame:
			cell = (Cell) model.getObject();
			mcellService.updateBorder(excelId,cell, step);
			break;
		case OperatorConstant.alignlevel:
			cell = (Cell) model.getObject();
			mcellService.updateAlignlevel(excelId,cell, step);
			break;
		case OperatorConstant.alignvertical:
			cell = (Cell) model.getObject();
			mcellService.updateAlignvertical(excelId,cell, step);
			break;
		case OperatorConstant.rowsinsert:
			RowOperate rowOperate = (RowOperate) model.getObject();
			mrowService.insertRow(excelId,rowOperate, step);
			break;
		case OperatorConstant.rowsdelete:
			RowOperate rowOperate2 = (RowOperate) model.getObject();
			mrowService.delRow(excelId,rowOperate2,step);
			break;
		case OperatorConstant.colsinsert:
			ColOperate colOperate = (ColOperate) model.getObject();
			
			/*Distribute dis = new Distribute();
			String sheetId = excelId+0;
			dis.addParam("colOperate", colOperate);
			dis.addParam("excelId", excelId);
			dis.addParam("step", step);
			dis.addParam("sheetId", sheetId);
			
			Target target1 = new Target(mcolService,"insertColSelf",1);
			Target target2 = new Target(mcolService,"insertColEffectMSheet",1);
			Target target3 = new Target(mcolService,"insertColEffectMCell",1);
			dis.add(target1);
			dis.add(target2);
			dis.add(target3);
			dis.exec();*/
			
			mcolService.insertCol(excelId,colOperate,  step);
			break;
		case OperatorConstant.colsdelete:
			ColOperate colOperate2 = (ColOperate) model.getObject();
			mcolService.delCol(excelId,colOperate2,  step);
			break;
		case OperatorConstant.paste:
			OuterPaste outerPaste = (OuterPaste) model.getObject();
			msheetService.paste(excelId,outerPaste, step);
			break;
		case OperatorConstant.copy:
			Copy copy = (Copy) model.getObject();
			msheetService.copy(excelId,copy,  step);
			break;
		case OperatorConstant.cut:
			Copy copy2 = (Copy) model.getObject();
			msheetService.cut( excelId,copy2, step);
			break;
		case OperatorConstant.frozen:
			Frozen frozen = (Frozen) model.getObject();
			msheetService.frozen(excelId,frozen,  step);
			break;
		case OperatorConstant.unFrozen:
			msheetService.unFrozen(excelId, step);
			break;
		case OperatorConstant.colswidth:
			ColWidth colWidth = (ColWidth) model.getObject();
			mcolService.updateColWidth(excelId,colWidth,  step);
			break;
		case OperatorConstant.colshide:
			ColOperate colHide = (ColOperate) model.getObject();
			mcolService.hideCol( excelId,colHide, step);
			break;
		case OperatorConstant.rowshide:
			RowOperate rowHide = (RowOperate) model.getObject();
			mrowService.hideRow(excelId,rowHide,  step);
			break;
		case OperatorConstant.colhideCancel:
			ColOperate colhideCancel = (ColOperate) model.getObject();
			mcolService.showCol( excelId,colhideCancel, step);
			break;
		case OperatorConstant.rowhideCancel:
			RowOperate rowhideCancel = (RowOperate) model.getObject();
			mrowService.showRow(excelId,rowhideCancel,  step);
			break;
		case OperatorConstant.rowsheight:
			RowHeight rowHeight = (RowHeight) model.getObject();
			mrowService.updateRowHeight( excelId,rowHeight, step);
			break;
		case OperatorConstant.expand:
			RowOrColExpand expand = (RowOrColExpand) model.getObject();
			String type = expand.getType();
			if ("row".equals(type)) {
				mrowService.addRow(excelId,expand.getNum(),  step);
			} else {
				mcolService.addCol(excelId,expand.getNum(),  step);
			}
			break;
		// case OperatorConstant.addRowLine:
		// RowLine rowLine = (RowLine) model.getObject();
		// int rowNum = rowLine.getNum();
		// sheetService.addRowLine(excelBook.getSheets().get(0),rowNum);
		// break;
		// case OperatorConstant.addColLine:
		// AddLine addLine = (AddLine) model.getObject();
		// int num = addLine.getNum();
		// sheetService.addColLine(excelBook.getSheets().get(0), num);
		// break;
		// case OperatorConstant.colorset:
		// cell = (Cell) model.getObject();
		// handleExcelService.colorSet(cell, excelBook);
	    case OperatorConstant.undo:
		    msheetService.undo(excelId);
	    break;
		case OperatorConstant.redo:
		    msheetService.redo(excelId);
		break;
		case OperatorConstant.cellLock:
			cell = (Cell) model.getObject();
			mcellService.updateLock(excelId,cell, step);
		break;
		// case OperatorConstant.batchcolorset:
		// ColorSet colorSet= (ColorSet) model.getObject();
		// handleExcelService.batchColorSet(colorSet, excelBook);
		default:
			break;
		}
		// System.out.println(JSON.toJSONString(versionHistory));
		// storeService.set(excelId,excelBook);
		System.out.println(step + "结束执行");
		logger.info("end excelId:" + excelId + ";step:" + step + ";reqPath:"
				+ reqPath);
		// mongodbServiceImpl.update(excelId, step, "step");
	}

	public int getStep() {
		return step;
	}

	public void setStep(int step) {
		this.step = step;
	}

	public String getKey() {
		return key;
	}

	public void setKey(String key) {
		this.key = key;
	}

	public Model getModel() {
		return model;
	}

	public void setModel(Model model) {
		this.model = model;
	}
  
}
