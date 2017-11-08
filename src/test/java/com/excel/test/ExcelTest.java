package com.excel.test;

import com.lorne.core.framework.utils.excel.model.LRow;
import com.lorne.core.framework.utils.excel.model.LSheet;
import com.lorne.core.framework.utils.excel.ExcelUtils;
import jxl.read.biff.BiffException;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * 测试类
 * 
 * @author yuliang
 * 
 */
public class ExcelTest {

	public static void main(String[] args) {
		List<LSheet> sheets = readExcel();
		writeExcel(sheets);

	}

	public static List<LSheet>  readExcel() {
		try {
			// 读取文件数据
			List<LSheet> sheets = ExcelUtils.getExcelData(new File("d://data.xls"));
			// 遍历打印
			for (LSheet s : sheets) {
				System.out
						.println("====================================================================================================");

				for (LRow row: s.getRows()) {
					for (String v:row.getContent()) {
						System.out.print(v + "\t\t");
					}
					System.out.println();
				}

				return sheets;
			}
		} catch (BiffException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

		return null;

	}

	public static void writeExcel(List<LSheet> sheets) {

		File file = new File("d://data_new.xls");
		try {
			 ExcelUtils.writeExcel(file,sheets);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}

	}

}
