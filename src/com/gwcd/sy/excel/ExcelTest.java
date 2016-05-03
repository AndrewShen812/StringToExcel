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
			// ����HSSFWorkbook����
			HSSFWorkbook wb = new HSSFWorkbook();
			// ����HSSFSheet����
			HSSFSheet sheet = wb.createSheet("sheet0");
			// ����HSSFRow����
			HSSFRow row = sheet.createRow(0);
			// ����HSSFCell����
			HSSFCell cell = row.createCell(0);
			// ���õ�Ԫ���ֵ
			cell.setCellValue("��Ԫ���е�����");
			// ���Excel�ļ�
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
			// ����HSSFWorkbook����(excel���ĵ�����)
			HSSFWorkbook wb = new HSSFWorkbook();
			// �����µ�sheet����excel�ı���
			HSSFSheet sheet = wb.createSheet("�ɼ���");
			// ��sheet�ﴴ����һ�У�����Ϊ������(excel����)��������0��65535֮����κ�һ��
			HSSFRow row1 = sheet.createRow(0);
			// ������Ԫ��excel�ĵ�Ԫ�񣬲���Ϊ��������������0��255֮����κ�һ��
			HSSFCell cell = row1.createCell(0);
			// ���õ�Ԫ������
			cell.setCellValue("ѧԱ���Գɼ�һ����");
			// �ϲ���Ԫ��CellRangeAddress����������α�ʾ��ʼ�У������У���ʼ�У� ������
			sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 3));
			// ��sheet�ﴴ���ڶ���
			HSSFRow row2 = sheet.createRow(1);
			// ������Ԫ�����õ�Ԫ������
			row2.createCell(0).setCellValue("����");
			row2.createCell(1).setCellValue("�༶");
			row2.createCell(2).setCellValue("���Գɼ�");
			row2.createCell(3).setCellValue("���Գɼ�");
			// ��sheet�ﴴ��������
			HSSFRow row3 = sheet.createRow(2);
			row3.createCell(0).setCellValue("����");
			row3.createCell(1).setCellValue("As178");
			row3.createCell(2).setCellValue(87);
			row3.createCell(3).setCellValue(78);
			
			HSSFRow row4 = sheet.createRow(3);
			row4.createCell(0).setCellValue("�ŷ�");
			row4.createCell(1).setCellValue("As255");
			row4.createCell(2).setCellValue(92);
			row4.createCell(3).setCellValue(83);

			// ���Excel�ļ�
			FileOutputStream output = new FileOutputStream("d:\\�ɼ���.xls");
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
