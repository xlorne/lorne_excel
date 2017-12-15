package com.lorne.core.framework.utils.excel;


import com.lorne.core.framework.utils.excel.model.LRow;
import com.lorne.core.framework.utils.excel.model.LSheet;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;


/**
 * Excel文件读取
 * @author yuliang
 *
 */

public class ExcelUtils {


	private static Cell withMergeCell(Sheet sheet,Cell cell,
			List<CellRangeAddress> ranges) {
		if(cell==null){
			return null;
		}
		if(ranges==null||ranges.size()==0){
			return cell;
		}
		for (CellRangeAddress range : ranges) {
			 if(range.isInRange(cell)){
				 int firstColumn = range.getFirstColumn();
				 int lastColumn = range.getLastColumn();
				 int firstRow = range.getFirstRow();
				 int lastRow = range.getLastRow();

				 Row row = sheet.getRow(firstRow);
				 Cell value = row.getCell(firstColumn);
				 return value;
			 }
		}
		return cell;
	}



	/**
	 * Excel文件读写
	 * @param filePath	Excel文件路径
	 * @return 所以的excel数据
	 * @throws Exception Exception
	 */
	public static List<LSheet> getExcelData(File filePath) throws Exception{
			Workbook wb = null;
			try {
				List<LSheet> sheets = new ArrayList<>();
				String fileName = filePath.getName();

				if (fileName.endsWith(".xlsx")) {
					wb = new XSSFWorkbook(filePath);
				} else {
					wb = new HSSFWorkbook(new FileInputStream(filePath));
				}
				int sheetNumbers = wb.getNumberOfSheets();

				for (int page = 0; page < sheetNumbers; page++) {
					Sheet sheet = wb.getSheetAt(page);

					LSheet lSheet = new LSheet();
					lSheet.setSheetName(sheet.getSheetName());

					List<CellRangeAddress> ranges = sheet.getMergedRegions();

					int rows = sheet.getPhysicalNumberOfRows();

					List<LRow> lRows = new ArrayList<>();

					for (int rowIndex = 0; rowIndex < rows; rowIndex++) {
						LRow lRow = new LRow();
						List<String> content = new ArrayList<>();
						Row row = sheet.getRow(rowIndex);
						if (row != null) {
							int colNumber = row.getPhysicalNumberOfCells();
							for (int colIndex = 0; colIndex < colNumber; colIndex++) {
								Cell val = withMergeCell(sheet, row.getCell(colIndex), ranges);
								String value = getCellValue(val);
								content.add(value);
							}
							lRow.setContent(content);
							lRows.add(lRow);
						}

					}
					lSheet.setRows(lRows);
					sheets.add(lSheet);

				}
				return sheets;
			}finally {
				if(wb!=null) {
					wb.close();
				}
			}


	}

	private static String getCellValue(Cell cell) 	{
		String cellvalue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			CellType cellType =  cell.getCellTypeEnum();
			switch (cellType) {
				// 如果当前Cell的Type为NUMERIC
				case NUMERIC: {
					short format = cell.getCellStyle().getDataFormat();
					if(format == 14 || format == 31 || format == 57 || format == 58){   //excel中的时间格式
						SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
						double value = cell.getNumericCellValue();
						Date date = DateUtil.getJavaDate(value);
						cellvalue = sdf.format(date);
					}
					// 判断当前的cell是否为Date
					else if (HSSFDateUtil.isCellDateFormatted(cell)) {  //先注释日期类型的转换，在实际测试中发现HSSFDateUtil.isCellDateFormatted(cell)只识别2014/02/02这种格式。
						// 如果是Date类型则，取得该Cell的Date值           // 对2014-02-02格式识别不出是日期格式
						Date date = cell.getDateCellValue();
						DateFormat formater = new SimpleDateFormat("yyyy-MM-dd");
						cellvalue= formater.format(date);
					} else { // 如果是纯数字
						// 取得当前Cell的数值
						cellvalue = NumberToTextConverter.toText(cell.getNumericCellValue());
					}
					break;
				}
				// 如果当前Cell的Type为STRIN
				case STRING: {
					// 取得当前的Cell字符串
					cellvalue = cell.getStringCellValue().replaceAll("'", "''");
					break;
				}
				case  BLANK: {
					cellvalue = "";
					break;
				}
				case  FORMULA: {
					try {
						DecimalFormat df = new DecimalFormat("#.##");    //格式化为四位小数，按自己需求选择；
						cellvalue = String.valueOf(df.format(cell.getNumericCellValue()));      //返回String类型
					} catch (Exception e) {
						cellvalue = String.valueOf(cell);
					}
					break;
				}
				// 默认的Cell值
				default:{
					cellvalue = " ";
				}
			}
		} else {
			cellvalue = "";
		}
		return cellvalue;
	}


	/**
	 * 写入excel文件
	 * @param file 文件路径
	 * @param sheets	数据
	 * @throws Exception Exception
	 */
	public static void writeExcel(File file,List<LSheet> sheets) throws Exception {
		Workbook wb = null;
		OutputStream os = null;
		try {
			// 创建工作薄
			String fileName = file.getName();
			os = new FileOutputStream(file);

			if (fileName.endsWith(".xlsx")) {
				wb = new XSSFWorkbook();
			} else {
				wb = new HSSFWorkbook();
			}
			// 创建新的一页
			for (int pageIndex = 0; pageIndex < sheets.size(); pageIndex++) {
				LSheet lSheet = sheets.get(pageIndex);
				Sheet sheet = wb.createSheet(lSheet.getSheetName());

				List<LRow> rows = lSheet.getRows();

				for (int row = 0; row < rows.size(); row++) {
					Row rowData = sheet.createRow(row);
					List<String> content = rows.get(row).getContent();
					for (int j = 0; j < content.size(); j++) {
						String v = content.get(j);
						rowData.createCell(j).setCellValue(v);
					}
				}
			}
			wb.write(os);

		}finally {
			if(wb!=null) {
				wb.close();
			}
			if(os!=null) {
				os.close();
			}
		}

	}

}
