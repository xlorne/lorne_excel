package com.lorne.core.framework.utils.excel;


import com.lorne.core.framework.utils.excel.model.LMergeCellModel;
import com.lorne.core.framework.utils.excel.model.LRow;
import com.lorne.core.framework.utils.excel.model.LSheet;
import jxl.Cell;
import jxl.Range;
import jxl.Sheet;
import jxl.Workbook;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;

import java.io.*;
import java.util.ArrayList;
import java.util.List;

/**
 * Excel文件读取
 * @author yuliang
 *
 */

public class ExcelUtils {

	/**
	 * 判断该单元格是否属于合并单元格
	 * @param x
	 * @param y
	 * @param lists
	 * @return 不存在与该单元格返回值为null
	 */
	private static String inMergeCells(int x, int y,
			List<LMergeCellModel> lists) {
		String value = null;
		if (lists == null || lists.size() <= 0) {
			return value;
		}
		for (LMergeCellModel e : lists) {
			if ((x >= e.getStartX() && x <= e.getEndX())
					&& (y >= e.getStartY() && y <= e.getEndY())) {
				value = e.getValue();
			}
		}
		return value;
	}

	/**
	 * Excel文件读写
	 * @param filePath	Excel文件路径
	 * @return 所以的excel数据
	 * @throws Exception Exception
	 */
	public static List<LSheet> getExcelData(File filePath) throws Exception {
		InputStream is = new FileInputStream(filePath);
		//获取存储Excel数据的Workbook对象
		Workbook rwb = Workbook.getWorkbook(is);
		//获取Excel的页数
		int slg = rwb.getSheets().length;

		List<LSheet> sheets = new ArrayList<>();
		//遍历Excel每页中的值
		for (int si = 0; si < slg; si++) {
			//获取当页的数据
			Sheet rs =  rwb.getSheet(si);

			LSheet lSheet = new LSheet();
			lSheet.setSheetName(rs.getName());

			//获取当前页中合并的单元格
			List<LMergeCellModel> lists = readRange(rs);
			//创建当页数据的数组对象
			int rows = rs.getRows();
			int columns = rs.getColumns();

			List<LRow> lRows = new ArrayList<>();
			//遍历获取该页下所有的单元格数据
			for (int i = 0; i < rows; i++) {
				LRow lRow = new LRow();
				List<String> content = new ArrayList<>();
				for (int j = 0; j < columns; j++) {
					//判断当前单元格是否存在与合并单元格中
					String mv = inMergeCells(j, i, lists);
					//获取当前单元格的值
					Cell v = rs.getCell(j, i);
					String value = mv == null ? v.getContents() : mv;
					content.add(value);
				}
				lRow.setContent(content);
				lRows.add(lRow);
			}

			lSheet.setRows(lRows);
			sheets.add(lSheet);
		}
		if(rwb!=null) {
			rwb.close();
		}
		if(is!=null){
			is.close();
		}
		return sheets;
	}


	private static List<LMergeCellModel> readRange(Sheet rs){
		Range[] ranges = rs.getMergedCells();
		List<LMergeCellModel> list = new ArrayList<LMergeCellModel>();
		//封装合并单元格的数据
		for (Range r : ranges) {
			LMergeCellModel cellModel = new LMergeCellModel();
			cellModel.setStartY(r.getTopLeft().getRow());
			cellModel.setEndY(r.getBottomRight().getRow());
			cellModel.setStartX(r.getTopLeft().getColumn());
			cellModel.setEndX(r.getBottomRight().getColumn());
			cellModel.setValue(r.getTopLeft().getContents());
			list.add(cellModel);
		}
		return list;
	}


	/**
	 * 写入excel文件
	 * @param file 文件路径
	 * @param sheets	数据
	 * @throws Exception Exception
	 */
	public static void writeExcel(File file,List<LSheet> sheets) throws Exception {
		OutputStream os = new FileOutputStream(file);
		// 创建工作薄
		WritableWorkbook workbook = Workbook.createWorkbook(os);
		// 创建新的一页
		for(int pageIndex=0;pageIndex<sheets.size();pageIndex++) {
			LSheet lSheet = sheets.get(pageIndex);
			WritableSheet sheet = workbook.createSheet(lSheet.getSheetName(),pageIndex);
			List<LRow> rows = lSheet.getRows();
			for (int row = 0; row < rows.size(); row++) {
				List<String> content = rows.get(row).getContent();
				for (int j = 0; j <content.size(); j++) {
					String v = content.get(j);
					Label l = new Label(j, row, v);
					sheet.addCell(l);
				}
			}
		}
		workbook.write();
		workbook.close();
		os.close();
	 
	}

}
