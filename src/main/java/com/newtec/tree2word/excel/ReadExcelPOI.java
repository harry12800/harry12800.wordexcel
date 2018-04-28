package com.newtec.tree2word.excel;

import java.io.FileInputStream;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ReadExcelPOI {

	public void readExcel() {
		try {

			String fileName = "E:\\xxx\\xxx\\a.xlsx";
			XSSFWorkbook xwb = new XSSFWorkbook(new FileInputStream(fileName));
			// 循环工作表Sheet
			for (int numSheet = 0; numSheet < xwb.getNumberOfSheets(); numSheet++) {
				XSSFSheet xSheet = xwb.getSheetAt(numSheet);
				if (xSheet == null) {
					continue;
				}

				// 循环行Row
				for (int rowNum = 0; rowNum <= xSheet.getLastRowNum(); rowNum++) {
					XSSFRow xRow = xSheet.getRow(rowNum);
					if (xRow == null) {
						continue;
					}

					// 循环列Cell
					for (int cellNum = 0; cellNum <= xRow.getLastCellNum(); cellNum++) {
						XSSFCell xCell = xRow.getCell(cellNum);
						if (xCell == null) {
							continue;
						}
						System.out.print("        " + getValue(xCell));
					}
					System.out.println();
				}
			}

		} catch (Exception e) {
			e.printStackTrace();
		}
	}

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
		ReadExcelPOI readExcelService = new ReadExcelPOI();
		readExcelService.readExcel();
	}
}