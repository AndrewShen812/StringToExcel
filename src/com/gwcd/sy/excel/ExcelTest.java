package com.gwcd.sy.excel;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;

public class ExcelTest {

	public static void main(String[] args) {
		ExcelTest excel = new ExcelTest();
		excel.CreateTest2();
	}

	public void createTest() {
		try {
			// 创建HSSFWorkbook对象
			HSSFWorkbook wb = new HSSFWorkbook();
			// 创建HSSFSheet对象
			HSSFSheet sheet = wb.createSheet("sheet0");
			// 创建HSSFRow对象
			HSSFRow row = sheet.createRow(0);
			// 创建HSSFCell对象
			HSSFCell cell = row.createCell(0);
			// 设置单元格的值
			cell.setCellValue("单元格中的中文");
			// 输出Excel文件
			FileOutputStream output = new FileOutputStream("d:\\workbook.xls");
			wb.write(output);
			output.flush();
			output.close();
			System.out.println("done.");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	public void CreateTest2() {
		try {
			// 创建HSSFWorkbook对象(excel的文档对象)
			HSSFWorkbook wb = new HSSFWorkbook();
			// 建立新的sheet对象（excel的表单）
			HSSFSheet sheet = wb.createSheet("成绩表");
			// 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
			HSSFRow row1 = sheet.createRow(0);
			// 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
			HSSFCell cell = row1.createCell(0);
			// 设置单元格内容
			cell.setCellValue("学员考试成绩一览表");
			// 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
			// 在sheet里创建第二行
			HSSFRow row2 = sheet.createRow(1);
			// 创建单元格并设置单元格内容
			row2.createCell(0).setCellValue("姓名");
			row2.createCell(1).setCellValue("班级");
			row2.createCell(2).setCellValue("笔试成绩");
			row2.createCell(3).setCellValue("机试成绩");
			// 在sheet里创建第三行
			HSSFRow row3 = sheet.createRow(2);
			row3.createCell(0).setCellValue("李明");
			row3.createCell(1).setCellValue("As178");
			row3.createCell(2).setCellValue(87);
			row3.createCell(3).setCellValue(78);
			
			HSSFRow row4 = sheet.createRow(3);
			row4.createCell(0).setCellValue("张飞");
			row4.createCell(1).setCellValue("As255");
			row4.createCell(2).setCellValue(92);
			row4.createCell(3).setCellValue(83);

			// 输出Excel文件
			FileOutputStream output = new FileOutputStream("d:\\成绩表.xls");
			wb.write(output);
			output.flush();
			output.close();
			System.out.println("test2 done.");
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}
