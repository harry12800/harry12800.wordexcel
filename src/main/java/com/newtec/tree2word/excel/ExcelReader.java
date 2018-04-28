package com.newtec.tree2word.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.harry12800.tools.FileUtils;

/**
 * 操作Excel表格的功能类
 * @param <SXSSFWorkbook>
 */
public class ExcelReader<SXSSFWorkbook> {
	static Excel2003 e2003 = new Excel2003();
	static Excel2007 e2007 = new Excel2007();

	public  static synchronized  boolean isExcel2003(InputStream is) {
		try {
			new HSSFWorkbook(is);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	public  static synchronized  boolean isExcel2007(InputStream is) {
		try {
			new XSSFWorkbook(is);
		} catch (Exception e) {
			return false;
		}
		return true;
	}

	public  static synchronized  boolean isExcel2003(String path) {
		FileInputStream is = null;
		try {
			is = new FileInputStream(path);
			new HSSFWorkbook(is);
			is.close();
		} catch (Exception e) {
			if (is != null)
				try {
					is.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			return false;
		}
		return true;
	}

	public  static synchronized   boolean isExcel2007(String path) {
		FileInputStream is = null;
		try {
			is = new FileInputStream(path);
			new XSSFWorkbook(is);
			is.close();
		} catch (Exception e) {
			if (is != null)
				try {
					is.close();
				} catch (IOException e1) {
					e1.printStackTrace();
				}
			return false;
		}
		return true;
	}

	public  static synchronized String[] readExcelTitle(String path) throws Exception {
		if (isExcel2003(path)) {
			//System.out.println("2003");
			return e2003.readExcelTitle(path);
		} else if (isExcel2007(path)) {
			//System.out.println("2007");
			return e2007.readExcelTitle(path);
		}
		throw new Exception("不是excel文件！");
	}

	public  static synchronized  Map<Integer, List<String>> readExcelContent(String path) throws Exception {
		if (isExcel2003(path)) {
			return e2003.readExcelContent(path);
		} else if (isExcel2007(path)) {
			return e2007.readExcelContent(path);
		}
		throw new Exception("不是excel文件！");
	}

	public static synchronized boolean writeCell(String path, Integer row, int column,
			String content) throws IOException {
		if (isExcel2003(path)) {
			return e2003.writeCell(path, row, column, content);
		} else if (isExcel2007(path)) {
			return e2007.writeCell(path, row, column, content);
		}
		return false;
	}

	public static synchronized boolean deleteRow(String path, int row) {
		if (isExcel2003(path)) {
			return e2003.deleteRow(path, row);
		} else if (isExcel2007(path)) {
			return e2007.deleteRow(path, row);
		}
		return false;
	}

	public static synchronized  boolean writeCells(String path, Set<ExcelPosition> set) throws Exception {
		if (isExcel2003(path)) {
			return e2003.writeCells(path, set);
		} else if (isExcel2007(path)) {
			return e2007.writeCells(path, set);
		}
		return false;
	}

	public static synchronized boolean writeCells(String srcpath, String desPath,
			Set<ExcelPosition> set) throws Exception {
		if(!new File(srcpath).exists()){
			throw new Exception("文件不存在！"+srcpath);
		}
		if (isExcel2003(srcpath)) {
			return e2003.writeCells(srcpath, desPath,set);
		} else if (isExcel2007(srcpath)) {
			return e2007.writeCells(srcpath, desPath,set);
		}
		return false;
	}
	/**
	 * 获取excel文件转化html的html文本内容。
	 * @param path
	 * @return
	 */
	public static synchronized String toHtml(String path){
		InputStream is = null;
		String htmlExcel = null;
		try {
			File sourcefile = new File(path);
			is = new FileInputStream(sourcefile);
			Workbook wb = WorkbookFactory.create(is);// 此WorkbookFactory在POI-3.10版本中使用需要添加dom4j
			if (wb instanceof XSSFWorkbook) {
				XSSFWorkbook xWb = (XSSFWorkbook) wb;
				htmlExcel = POIReadExcelToHtml.getExcelInfo(xWb, true);
			} else if (wb instanceof HSSFWorkbook) {
				HSSFWorkbook hWb = (HSSFWorkbook) wb;
				htmlExcel = POIReadExcelToHtml.getExcelInfo(hWb, true);
			}
			System.out.println(htmlExcel);
			//FileUtils.writeContent("D:\\1.html", htmlExcel);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			try {
				is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return htmlExcel;
	}
	/**
	 * 将excel文件生成html保存到文件htmlpath中
	 * @param path  需要转化的excel文件路径
	 * @param htmlPath 生成的html路径
	 * @return htmlPath
	 * @throws Exception 
	 */
	public static synchronized String toHtml(String path,String htmlPath) throws Exception{
		InputStream is = null;
		String htmlExcel = null;
		try {
			File sourcefile = new File(path);
			is = new FileInputStream(sourcefile);
			Workbook wb = WorkbookFactory.create(is);// 此WorkbookFactory在POI-3.10版本中使用需要添加dom4j
			if (wb instanceof XSSFWorkbook) {
				XSSFWorkbook xWb = (XSSFWorkbook) wb;
				htmlExcel = POIReadExcelToHtml.getExcelInfo(xWb, true);
			} else if (wb instanceof HSSFWorkbook) {
				HSSFWorkbook hWb = (HSSFWorkbook) wb;
				htmlExcel = POIReadExcelToHtml.getExcelInfo(hWb, true);
			}
			System.out.println(htmlExcel);
			htmlExcel = "<meta http-equiv='Content-Type' content='text/html; charset=utf-8' />"+htmlExcel;
			FileUtils.writeContent(htmlPath, htmlExcel);
		} catch (Exception e) {
			 e.printStackTrace();
			 throw e;
		} finally {
			try {
				if(is!=null)
					is.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return htmlPath;
	}
	/**
	 * 工作薄对象
	 */
	private static org.apache.poi.xssf.streaming.SXSSFWorkbook wb;
	
	/**
	 * 工作表对象
	 */
	private static Sheet sheet;
	
	/**
	 * 样式列表
	 */
	private static Map<String, CellStyle> styles;

	public static synchronized void createExcel(String path,
			Set<ExcelPosition> set) {
		wb = new org.apache.poi.xssf.streaming.SXSSFWorkbook(500);
		styles = createStyles(wb);
		sheet = wb.createSheet("export");
		sheet.autoSizeColumn((short)0); //调整第一列宽度
        sheet.autoSizeColumn((short)1); //调整第二列宽度
        sheet.autoSizeColumn((short)2); //调整第三列宽度
        sheet.autoSizeColumn((short)3); //调整第四列宽度
		for (ExcelPosition ep : set) {
			Row cellRow = sheet.getRow(ep.getRow());
			if (cellRow == null) {
				cellRow = sheet.createRow(ep.getRow());
			}
			cellRow.setHeightInPoints(30);
			Cell cell = cellRow.getCell(ep.getCol());
			if (cell == null) {
				cell = cellRow.createCell(ep.getCol());
			}
			if (ep.getRow() == 0)
				cell.setCellStyle(styles.get("title"));
			else {
				cell.setCellStyle(styles.get("data2"));
			}
			cell.setCellValue(ep.getContent());
		}

		FileOutputStream out = null;
		try {
			out = new FileOutputStream(path);
			wb.write(out);
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			try {
				if (out != null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
	}

	private static Map<String, CellStyle> createStyles(
			org.apache.poi.xssf.streaming.SXSSFWorkbook wb) {
		Map<String, CellStyle> styles = new HashMap<String, CellStyle>();
		
		CellStyle style = wb.createCellStyle();
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		Font titleFont = wb.createFont();
		titleFont.setFontName("Arial");
		titleFont.setFontHeightInPoints((short) 16);
		titleFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		style.setFont(titleFont);
		styles.put("title", style);

		style = wb.createCellStyle();
		style.setVerticalAlignment(CellStyle.VERTICAL_CENTER);
		style.setBorderRight(CellStyle.BORDER_THIN);
		style.setRightBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderLeft(CellStyle.BORDER_THIN);
		style.setLeftBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderTop(CellStyle.BORDER_THIN);
		style.setTopBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setBorderBottom(CellStyle.BORDER_THIN);
		style.setBottomBorderColor(IndexedColors.GREY_50_PERCENT.getIndex());
		Font dataFont = wb.createFont();
		dataFont.setFontName("Arial");
		dataFont.setFontHeightInPoints((short) 10);
		style.setFont(dataFont);
		styles.put("data", style);
		
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_LEFT);
		styles.put("data1", style);

		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_CENTER);
		styles.put("data2", style);

		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
		style.setAlignment(CellStyle.ALIGN_RIGHT);
		styles.put("data3", style);
		
		style = wb.createCellStyle();
		style.cloneStyleFrom(styles.get("data"));
//		style.setWrapText(true);
		style.setAlignment(CellStyle.ALIGN_CENTER);
		style.setFillForegroundColor(IndexedColors.GREY_50_PERCENT.getIndex());
		style.setFillPattern(CellStyle.SOLID_FOREGROUND);
		Font headerFont = wb.createFont();
		headerFont.setFontName("Arial");
		headerFont.setFontHeightInPoints((short) 10);
		headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
		headerFont.setColor(IndexedColors.WHITE.getIndex());
		style.setFont(headerFont);
		styles.put("header", style);
		
		return styles;
	}
}