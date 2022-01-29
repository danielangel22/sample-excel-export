package com.example.excelexport.util;

import java.io.IOException;
import java.util.List;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Component;

import com.example.excelexport.User;

@Component
public class ExcelGenerator {

	private XSSFWorkbook workbook;
	private XSSFSheet sheet;

	private void writeHeaderLine(Class<?> classs) {
		workbook = new XSSFWorkbook();
		sheet = workbook.createSheet("excel-export");
		Row row = sheet.createRow(0);
		var style = workbook.createCellStyle();
		var font = workbook.createFont();
		font.setBold(true);
		font.setFontHeight(16);
		style.setFont(font);
		var fields = classs.getDeclaredFields();

		for (var i = 0; i < fields.length; i++) {
			if (!fields[i].getName().contentEquals("serialVersionUID")) {
				System.out.println("atributo: " + fields[i].getName());
				createCell(row, i-1, fields[i].getName().toUpperCase(), style);
			}
		}
	}

	private void createCell(Row row, int columnCount, Object value, CellStyle style) {
		sheet.autoSizeColumn(columnCount);
		Cell cell = row.createCell(columnCount);
		if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else {
			cell.setCellValue((String) value);
		}
		cell.setCellStyle(style);
	}

	private void writeDataLines(List<User> results) {
		int rowCount = 1;

		CellStyle style = workbook.createCellStyle();
		XSSFFont font = workbook.createFont();
		font.setFontHeight(14);
		style.setFont(font);
		for (User user : results) {
			Row row = sheet.createRow(rowCount++);
			createCell(row, 0, user.getId(), style);
			createCell(row, 1, user.getEmail(), style);
			createCell(row, 2, user.getFullName(), style);
			createCell(row, 3, user.getPassword(), style);
			createCell(row, 4, user.isEnabled(), style);

		}
	}

	public void export(HttpServletResponse response, List<User> r) throws IOException {
		writeHeaderLine(User.class);
		writeDataLines(r);

		ServletOutputStream outputStream = response.getOutputStream();
		workbook.write(outputStream);
		workbook.close();

		outputStream.close();
	}

}
