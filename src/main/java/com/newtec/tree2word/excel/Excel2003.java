package com.newtec.tree2word.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

public class Excel2003 implements IExcelOperate {

	private POIFSFileSystem fs;
	private HSSFWorkbook wb;
	private HSSFSheet sheet;
	private HSSFRow row;

	/**
	 * 读取Excel表格表头的内容
	 * 
	 * @param InputStream
	 * @return String 表头内容的数组
	 */
	public String[] readExcelTitle(String path) {
		try {
			InputStream is = new FileInputStream(path);
			fs = new POIFSFileSystem(is);
			wb = new HSSFWorkbook(fs);
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet = wb.getSheetAt(0);
		row = sheet.getRow(0);
		if(row==null)return new String[]{};
		// 标题总列数
		int colNum = row.getPhysicalNumberOfCells();
		System.out.println("colNum:" + colNum);
		String[] title = new String[colNum];
		for (int i = 0; i < colNum; i++) {
			// title[i] = getStringCellValue(row.getCell((short) i));
			title[i] = getCellFormatValue(row.getCell(i));
		}
		return title;
	}

	/**
	 * 读取Excel数据内容
	 * 
	 * @param InputStream
	 * @return Map 包含单元格数据内容的Map对象
	 */
	public Map<Integer, List<String>> readExcelContent(String path) {
		
		Map<Integer, List<String>> values = new LinkedHashMap<Integer, List<String>>();
		
		String str = "";
		try {
			InputStream is = new FileInputStream(path);
			fs = new POIFSFileSystem(is);
			wb = new HSSFWorkbook(fs);
		} catch (IOException e) {
			e.printStackTrace();
		}
		sheet = wb.getSheetAt(0);
		// 得到总行数
		int rowNum = sheet.getLastRowNum();
		//System.out.println("ROW:" + rowNum);
		row = sheet.getRow(0);
		int colNum = row.getPhysicalNumberOfCells();
		// 正文内容应该从第二行开始,第一行为表头的标题

		for (int i = 1; i <= rowNum; i++) {
			row = sheet.getRow(i);
			if (row == null) {
				values.put(i, null);
				continue;
			}
			List<String> list = new ArrayList<String>();
			int j = 0;
			while (j < colNum) {
				// 每个单元格的数据内容用"-"分割开，以后需要时用String类的replace()方法还原数据
				// 也可以将每个单元格的数据设置到一个javabean的属性中，此时需要新建一个javabean
				// str += getStringCellValue(row.getCell((short) j)).trim() +
				// "-";
				Object o = row.getCell(j);
				if (o == null) {
					list.add("");
				} else {
					str = getCellFormatValue(row.getCell(j)).trim();
					list.add(str);
				}
				j++;
			}
			values.put(i, list);
		}
		return values;
	}

	/**
	 * 获取单元格数据内容为字符串类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	private String getStringCellValue(HSSFCell cell) {
		String strCell = "";
		switch (cell.getCellType()) {
		case HSSFCell.CELL_TYPE_STRING:
			strCell = cell.getStringCellValue();
			break;
		case HSSFCell.CELL_TYPE_NUMERIC:
			strCell = String.valueOf(cell.getNumericCellValue());
			break;
		case HSSFCell.CELL_TYPE_BOOLEAN:
			strCell = String.valueOf(cell.getBooleanCellValue());
			break;
		case HSSFCell.CELL_TYPE_BLANK:
			strCell = "";
			break;
		default:
			strCell = "";
			break;
		}
		if (strCell.equals("") || strCell == null) {
			return "";
		}
		return strCell;
	}

	/**
	 * 获取单元格数据内容为日期类型的数据
	 * 
	 * @param cell
	 *            Excel单元格
	 * @return String 单元格数据内容
	 */
	@SuppressWarnings({ "deprecation", "unused" })
	private String getDateCellValue(HSSFCell cell) {
		String result = "";
		try {
			int cellType = cell.getCellType();
			if (cellType == HSSFCell.CELL_TYPE_NUMERIC) {
				Date date = cell.getDateCellValue();
				result = (date.getYear() + 1900) + "-" + (date.getMonth() + 1) + "-" + date.getDate();
			} else if (cellType == HSSFCell.CELL_TYPE_STRING) {
				String date = getStringCellValue(cell);
				result = date.replaceAll("[年月]", "-").replace("日", "").trim();
			} else if (cellType == HSSFCell.CELL_TYPE_BLANK) {
				result = "";
			}
		} catch (Exception e) {
			System.out.println("日期格式不正确!");
			e.printStackTrace();
		}
		return result;
	}

	public static void main(String[] args) {
		new Excel2003().readExcelContent("C:\\Users\\Administrator\\Desktop\\1.xls");
	}

	/**
	 * 根据HSSFCell类型设置数据
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellFormatValue(HSSFCell cell) {
		String cellvalue = "";
		if (cell != null) {
			switch (cell.getCellType()) {// 判断当前Cell的Type
			case HSSFCell.CELL_TYPE_NUMERIC:// 如果当前Cell的Type为NUMERIC
			case HSSFCell.CELL_TYPE_FORMULA: {
				// 判断当前的cell是否为Date
				if (HSSFDateUtil.isCellDateFormatted(cell)) {
					// 如果是Date类型则，转化为Data格式
					// 方法1：这样子的data格式是带时分秒的：2011-10-12 0:00:00
					// cellvalue = cell.getDateCellValue().toLocaleString();
					// 方法2：这样子的data格式是不带带时分秒的：2011-10-12
					Date date = cell.getDateCellValue();
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellvalue = sdf.format(date);
				}
				// 如果是纯数字
				else {
					// 取得当前Cell的数值
					DecimalFormat df = new DecimalFormat("0");
					String whatYourWant = df.format(cell.getNumericCellValue());
					cellvalue = String.valueOf(cell.getNumericCellValue());
					cellvalue = whatYourWant;
				}
				break;
			}
			case HSSFCell.CELL_TYPE_STRING:// 如果当前Cell的Type为STRIN
				cellvalue = cell.getRichStringCellValue().getString();// 取得当前的Cell字符串
				break;
			default:	// 默认的Cell值
				cellvalue = " ";
			}
		}
		return cellvalue;
	}

	public boolean writeCell(String path, Integer row, int changeIndex, String content) throws IOException {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(path));
		HSSFSheet sheet = workbook.getSheetAt(0);

		HSSFRow rowCell = sheet.getRow(row);
		if (rowCell == null) {
			rowCell = sheet.createRow(row);
		}
		HSSFCell cell = rowCell.getCell(changeIndex);
		if (cell == null) {
			cell = rowCell.createCell(changeIndex);
		}
		cell.setCellValue(content);
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(path);
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if(out!=null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}

		return true;
	}

	@Override
	public boolean deleteRow(String path, int row) {
		// TODO Auto-generated method stub
		return true;
	}

	@Override
	public boolean writeCells(String path, Set<ExcelPosition> set) throws Exception {
		HSSFWorkbook workbook = null;
		if(new File(path).exists())
			  workbook = new HSSFWorkbook(new FileInputStream(path));
		else{
			  workbook = new HSSFWorkbook();
		}
		HSSFSheet sheet =null;
		try{
			sheet = workbook.getSheetAt(0);
		}catch (Exception e) {
			sheet = workbook.createSheet();
		}
		if(sheet==null)
		{
			sheet = workbook.createSheet();
		}
		for (ExcelPosition ep : set) {
			HSSFRow rowCell = sheet.getRow(ep.getRow());
			if (rowCell == null) {
				rowCell = sheet.createRow(ep.getRow());
			}
			HSSFCell cell = rowCell.getCell(ep.getCol());
			if (cell == null) {
				cell = rowCell.createCell(ep.getCol());
			}
			cell.setCellValue(ep.getContent());
		}
		try (FileOutputStream out = new FileOutputStream(path);){
			workbook.write(out);
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		}
		return true;
	}

	@Override
	public boolean writeCells(String srcpath, String descPath, Set<ExcelPosition> set) throws Exception {
		HSSFWorkbook workbook = new HSSFWorkbook(new FileInputStream(srcpath));
		HSSFSheet sheet = workbook.getSheetAt(0);
		for (ExcelPosition ep : set) {
			HSSFRow rowCell = sheet.getRow(ep.getRow());
			if (rowCell == null) {
				rowCell = sheet.createRow(ep.getRow());
			}
			HSSFCell cell = rowCell.getCell(ep.getCol());
			if (cell == null) {
				cell = rowCell.createCell(ep.getCol());
			}
			cell.setCellValue(ep.getContent());
		}
		FileOutputStream out = null;
		try {
			out = new FileOutputStream(descPath);
			workbook.write(out);
		} catch (IOException e) {
			System.out.println("原来："+e.getMessage());
			throw e;
		} finally {
			try {
				if(out!=null)
					out.close();
			} catch (IOException e) {
				e.printStackTrace();
			}
		}
		return true;
	}

	@Override
	public String toHtml(String path) {
		// TODO Auto-generated method stub
		return null;
	}

}
