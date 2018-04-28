package com.newtec.tree2word.excel;

import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class WriteExcelPOI {

	public void writeExcelPOI() {
		try {

			String fileName = "E:\\xxx\\xxx\\a.xlsx";
			XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(fileName));

			XSSFSheet xSheet = xwb.getSheetAt(0);
			XSSFRow xRow = xSheet.createRow(0);
			XSSFCell xCell = xRow.createCell(0);
			xCell.setCellValue("asdfasd");
			FileOutputStream out = new FileOutputStream(fileName);
			xwb.write(out);
			out.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	@SuppressWarnings("unused")
	private String getValue(XSSFCell xCell) {
		if (xCell.getCellType() == XSSFCell.CELL_TYPE_BOOLEAN) {
			return String.valueOf(xCell.getBooleanCellValue());
		} else if (xCell.getCellType() == XSSFCell.CELL_TYPE_NUMERIC) {
			return String.valueOf(xCell.getNumericCellValue());
		} else {
			return String.valueOf(xCell.getStringCellValue());
		}

	}

	public static void main(String[] args) {
		WriteExcelPOI a = new WriteExcelPOI();
		a.writeExcelPOI();
	}
}

