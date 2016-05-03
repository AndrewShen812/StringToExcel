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
			// step 1: 获得SAX解析器工厂实例
			SAXParserFactory factory = SAXParserFactory.newInstance();

			// step 2: 获得SAX解析器实例
			SAXParser parser = factory.newSAXParser();

			// step 3: 开始进行解析
			// 传入待解析的文档的处理器
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
		// 创建HSSFWorkbook对象(excel的文档对象)
		HSSFWorkbook wb = new HSSFWorkbook();
		// 建立新的sheet对象（excel的表单）
		HSSFSheet sheet = wb.createSheet("string对照");
		// 在sheet里创建第一行，参数为行索引(excel的行)，可以是0～65535之间的任何一个
		HSSFRow rowTitle = sheet.createRow(0);
		// 创建单元格（excel的单元格，参数为列索引，可以是0～255之间的任何一个
		HSSFCell cell = rowTitle.createCell(0);
		// 设置单元格内容
		cell.setCellValue("string多语言对照表");
		// 合并单元格CellRangeAddress构造参数依次表示起始行，截至行，起始列， 截至列
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 2));
		HSSFRow rowEvent = sheet.createRow(1);
		rowEvent.createCell(0).setCellValue("string_id");
		rowEvent.createCell(1).setCellValue("中文zh");
		rowEvent.createCell(2).setCellValue("英文en");
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
		// 输出Excel文件
		FileOutputStream output = new FileOutputStream("d:\\string对照表.xls");
		wb.write(output);
		output.flush();
		output.close();
	}
}
