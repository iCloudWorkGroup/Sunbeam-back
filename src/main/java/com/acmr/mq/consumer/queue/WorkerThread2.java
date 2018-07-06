package com.acmr.mq.consumer.queue;

import org.apache.log4j.Logger;

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

	public WorkerThread2(int step, String key, Model model,
			MCellService mcellService, MRowService mrowService,
			MColService mcolService, MSheetService msheetService) {

		this.step = step;
		this.key = key;
		this.model = model;
		this.mcellService = mcellService;
		this.mrowService = mrowService;
		this.mcolService = mcolService;
		this.msheetService = msheetService;
	}

	@Override
	public void run() {
		while (true) {
			int memStep = msheetService.getStep(key, key + 0);
			if (memStep + 1 == step) {
				System.out.println(step + "开始执行");
				logger.info("**********begin excelId : " + model.getExcelId()
						+ " === step : " + step + "== reqPath : "
						+ model.getReqPath());
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

	private void handleMessage(Model model) {
		int reqPath = model.getReqPath();
		String excelId = model.getExcelId();
		int step = model.getStep();

		Cell cell = null;
		switch (reqPath) {
		case OperatorConstant.textData:
			CellContent cellContent = (CellContent) model.getObject();
			mcellService.saveContent(cellContent, step, excelId);
			break;
		case OperatorConstant.fontsize:
			cell = (Cell) model.getObject();
			mcellService.updateFontSize(cell, step, excelId);
			break;
		case OperatorConstant.fontfamily:
			cell = (Cell) model.getObject();
			mcellService.updateFontFamily(cell, step, excelId);
			break;
		case OperatorConstant.fontweight:
			cell = (Cell) model.getObject();
			mcellService.updateFontWeight(cell, step, excelId);
			break;
		case OperatorConstant.fontitalic:
			cell = (Cell) model.getObject();
			mcellService.updateFontItalic(cell, step, excelId);
			break;
		case OperatorConstant.fontcolor:
			cell = (Cell) model.getObject();
			mcellService.updateFontColor(cell, step, excelId);
			break;
		case OperatorConstant.wordWrap:
			cell = (Cell) model.getObject();
			mcellService.updateWordwrap(cell, step, excelId);
			break;

		case OperatorConstant.fillbgcolor:
			cell = (Cell) model.getObject();
			mcellService.updateBgColor(cell, step, excelId);
			break;
		case OperatorConstant.textDataformat:
		    CellFormate cellFormate = (CellFormate) model.getObject();
	        mcellService.updateFormat(cellFormate, step, excelId);
		 break;
		//
		// case OperatorConstant.commentset:
		// Comment comment = (Comment) model.getObject();
		// handleExcelService.setComment(excelBook,
		// comment,versionHistory,step);
		// storeService.set(excelId+"_history", versionHistory);
		// break;
		case OperatorConstant.merge:
			cell = (Cell) model.getObject();
			mcellService.mergeCell(cell, step, excelId);
			break;
		case OperatorConstant.mergedelete:
			cell = (Cell) model.getObject();
			mcellService.splitCell(cell, step, excelId);
			break;
		case OperatorConstant.frame:
			cell = (Cell) model.getObject();
		    mcellService.updateBorder(cell, step, excelId);
			break;
		case OperatorConstant.alignlevel:
			cell = (Cell) model.getObject();
			mcellService.updateAlignlevel(cell, step, excelId);
			break;
		case OperatorConstant.alignvertical:
			cell = (Cell) model.getObject();
			mcellService.updateAlignvertical(cell, step, excelId);
		case OperatorConstant.rowsinsert:
			RowOperate rowOperate = (RowOperate) model.getObject();
			mrowService.insertRow(rowOperate, excelId, step);
			break;
		case OperatorConstant.rowsdelete:
			RowOperate rowOperate2 = (RowOperate) model.getObject();
			mrowService.delRow(rowOperate2, excelId, step);
			break;
		case OperatorConstant.colsinsert:
			ColOperate colOperate = (ColOperate) model.getObject();
			mcolService.insertCol(colOperate, excelId, step);
			break;
		case OperatorConstant.colsdelete:
			ColOperate colOperate2 = (ColOperate) model.getObject();
			mcolService.delCol(colOperate2, excelId, step);
			break;
		case OperatorConstant.paste:
			OuterPaste outerPaste = (OuterPaste) model.getObject();
			msheetService.paste(outerPaste, excelId, step);
			break;
		case OperatorConstant.copy:
			Copy copy = (Copy) model.getObject();
			msheetService.copy(copy, excelId, step);
			break;
		case OperatorConstant.cut:
			Copy copy2 = (Copy) model.getObject();
			msheetService.cut(copy2, excelId, step);
			break;
		case OperatorConstant.frozen:
			Frozen frozen = (Frozen) model.getObject();
			msheetService.frozen(frozen, excelId, step);
			break;
		case OperatorConstant.unFrozen:
			msheetService.unFrozen(excelId, step);
			break;
		case OperatorConstant.colswidth:
			ColWidth colWidth = (ColWidth) model.getObject();
			mcolService.updateColWidth(colWidth, excelId, step);
			break;
		case OperatorConstant.colshide:
			ColOperate colHide = (ColOperate) model.getObject();
			mcolService.hideCol(colHide, excelId, step);
			break;
		case OperatorConstant.rowshide:
			RowOperate rowHide = (RowOperate) model.getObject();
			mrowService.hideRow(rowHide, excelId, step);
			break;
		case OperatorConstant.colhideCancel:
			ColOperate colhideCancel = (ColOperate) model.getObject();
			mcolService.showCol(colhideCancel, excelId, step);
			break;
		case OperatorConstant.rowhideCancel:
			RowOperate rowhideCancel = (RowOperate) model.getObject();
			mrowService.showRow(rowhideCancel, excelId, step);
			break;
		case OperatorConstant.rowsheight:
			RowHeight rowHeight = (RowHeight) model.getObject();
			mrowService.updateRowHeight(rowHeight, excelId, step);
			break;
		case OperatorConstant.expand:
			RowOrColExpand expand = (RowOrColExpand) model.getObject();
			String type = expand.getType();
			if("row".equals(type)){
				 mrowService.addRow(expand.getNum(), excelId, step);
			}else{
				 mcolService.addCol(expand.getNum(), excelId, step);
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
		// case OperatorConstant.undo:
		// sheetService.undo(versionHistory, step,
		// excelBook.getSheets().get(0));
		// storeService.set(excelId+"_history", versionHistory);
		// break;
		// case OperatorConstant.redo:
		// sheetService.redo(versionHistory, step,
		// excelBook.getSheets().get(0));
		// storeService.set(excelId+"_history", versionHistory);
		// break;
		// case OperatorConstant.batchcolorset:
		// ColorSet colorSet= (ColorSet) model.getObject();
		// handleExcelService.batchColorSet(colorSet, excelBook);
		default:
			break;
		}
		// System.out.println(JSON.toJSONString(versionHistory));
		// storeService.set(excelId,excelBook);
		System.out.println(step + "结束执行");
		logger.info("**********end excelId : " + excelId + " === step : " + step
				+ "== reqPath : " + reqPath);
		// mongodbServiceImpl.update(excelId, step, "step");
	}

}
