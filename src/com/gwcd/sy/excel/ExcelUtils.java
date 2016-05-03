package com.gwcd.sy.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;

import javax.xml.parsers.ParserConfigurationException;
import javax.xml.parsers.SAXParser;
import javax.xml.parsers.SAXParserFactory;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.xml.sax.SAXException;

public class ExcelUtils implements IParseListener {

	private static ExcelUtils mInstance;

	private ExcelUtils() {
	}

	public static ExcelUtils getInstance() {
		if (mInstance == null) {
			mInstance = new ExcelUtils();
		}

		return mInstance;
	}

	public void doWiriteExcel(String xmlPath, String xlsPath) {
		try {
			// step 1: ���SAX����������ʵ��
			SAXParserFactory factory = SAXParserFactory.newInstance();

			// step 2: ���SAX������ʵ��
			SAXParser parser = factory.newSAXParser();

			// step 3: ��ʼ���н���
			// ������������ĵ��Ĵ�����
			ParseHandler handler = new ParseHandler();
			handler.setParseListener(this);
			parser.parse(new File(xmlPath), handler);
		} catch (SAXException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} catch (ParserConfigurationException e) {
			e.printStackTrace();
		}
	}

	@Override
	public void onStart() {
		System.out.println("onStart parse xml file");
	}

	@Override
	public void onFinish(Map<String, String> stringMap) {
		System.out.println("onFinish parse xml file");
		try {
			writeToExcel(stringMap);
		} catch (IOException e) {
			e.printStackTrace();
		}
	}

	private void writeToExcel(Map<String, String> stringMap) throws IOException {
		if (stringMap == null || stringMap.isEmpty()) {
			return;
		}
		// ����HSSFWorkbook����(excel���ĵ�����)
		HSSFWorkbook wb = new HSSFWorkbook();
		// �����µ�sheet����excel�ı���
		HSSFSheet sheet = wb.createSheet("string����");
		// ��sheet�ﴴ����һ�У�����Ϊ������(excel����)��������0��65535֮����κ�һ��
		HSSFRow rowTitle = sheet.createRow(0);
		// ������Ԫ��excel�ĵ�Ԫ�񣬲���Ϊ��������������0��255֮����κ�һ��
		HSSFCell cell = rowTitle.createCell(0);
		// ���õ�Ԫ������
		cell.setCellValue("string�����Զ��ձ�");
		// �ϲ���Ԫ��CellRangeAddress����������α�ʾ��ʼ�У������У���ʼ�У� ������
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		HSSFRow rowEvent = sheet.createRow(1);
		rowEvent.createCell(0).setCellValue("string_id");
		rowEvent.createCell(1).setCellValue("����zh");
		rowEvent.createCell(2).setCellValue("Ӣ��en");
		HSSFRow rowValue = null;
		Set<String> keys = stringMap.keySet();
		int rownum = 0;
		for (String key : keys) {
			String value = stringMap.get(key);
			rowValue = sheet.createRow(rownum + 2);
			rowValue.createCell(0).setCellValue(key);
			rowValue.createCell(1).setCellValue(value);
			rownum++;
		}
		// ���Excel�ļ�
		FileOutputStream output = new FileOutputStream("d:\\string���ձ�.xls");
		wb.write(output);
		output.flush();
		output.close();
	}
}
