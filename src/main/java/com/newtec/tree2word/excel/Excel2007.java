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

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import cn.harry12800.tools.FileUtils;
 

public class Excel2007 implements IExcelOperate{

	@Override
	public String[] readExcelTitle(String path) {
		try {
			InputStream in = new FileInputStream(path);
			XSSFWorkbook xwb = new XSSFWorkbook(in);
			XSSFSheet xSheet = xwb.getSheetAt(0);
			if (xSheet == null) {
				return null;
			}
			// 循环行Row
			XSSFRow xRow = xSheet.getRow(0);
			if (xRow == null) {
				return null;
			}
			String[] title = new String[xRow.getLastCellNum()];
			// 循环列Cell
			for (int cellNum = 0; cellNum <= xRow.getLastCellNum(); cellNum++) {
				XSSFCell xCell = xRow.getCell(cellNum);
				if (xCell == null) {
					continue;
				}
				title[cellNum] = getCellFormatValue(xCell);
			}
			return title;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	 
	@Override
	public Map<Integer, List<String>> readExcelContent(String path) {
		Map<Integer, List<String>> values = new LinkedHashMap<Integer, List<String>>();
		try {
			InputStream in = new FileInputStream(path);
			XSSFWorkbook xwb = new XSSFWorkbook(in);
			XSSFSheet xSheet = xwb.getSheetAt(0);
			if (xSheet == null) {
				return null;
			}
			// 循环行Row
			for (int rowNum = 1; rowNum <= xSheet.getLastRowNum(); rowNum++) {
				XSSFRow xRow = xSheet.getRow(rowNum);
				if (xRow == null) {
					continue;
				}
				List<String> title = new ArrayList<String>(0) ;
				// 循环列Cell
				for (int cellNum = 0; cellNum <= xRow.getLastCellNum(); cellNum++) {
					XSSFCell xCell = xRow.getCell(cellNum);
					if (xCell == null) {
						title.add("");
						continue;
					}
					title.add(getCellFormatValue(xCell));
				}
				values.put(rowNum, title);
			}
			return values;
		} catch (Exception e) {
			e.printStackTrace();
		}
		return null;
	}
	/**
	 * 根据HSSFCell类型设置数据
	 * 
	 * @param cell
	 * @return
	 */
	private String getCellFormatValue(XSSFCell   cell) {
		String cellvalue = "";
		if (cell != null) {
			// 判断当前Cell的Type
			switch (cell.getCellType()) {
			// 如果当前Cell的Type为NUMERIC
			case XSSFCell.CELL_TYPE_NUMERIC:
			case XSSFCell.CELL_TYPE_FORMULA: {
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
				/*
				 * 
				 * id ,编码,描述，订单日期第一列，数量，尺寸，材质，刀模号，非零号，颜色，备注，导入日期，修改日期。
				 */
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
			case XSSFCell.CELL_TYPE_BOOLEAN:
				cellvalue = String.valueOf(cell.getBooleanCellValue());
				break;
			// 如果当前Cell的Type为STRIN
			case XSSFCell.CELL_TYPE_STRING:
				// 取得当前的Cell字符串
				cellvalue = cell.getRichStringCellValue().getString();
				break;
			// 默认的Cell值
			default:
				cellvalue = " ";
			}
		} else {
			cellvalue = "";
		}
		return cellvalue;

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

	@Override
	public boolean writeCell(String path, Integer row, int changeIndex,
			String content) throws IOException {
		InputStream in = new FileInputStream(path);
		XSSFWorkbook xwb = new XSSFWorkbook(in);
		XSSFSheet xSheet = xwb.getSheetAt(0);
		XSSFRow cellRow = xSheet.getRow(row);
		if(cellRow ==null){
			cellRow = xSheet.createRow(row);
		}
		XSSFCell cell = cellRow.getCell(changeIndex);
		if(cell == null) {
			cell = cellRow.createCell(changeIndex);
		}
		File file = new File(path);
		String temp = file.getParentFile().getAbsoluteFile()+"\\tmp.xlsx";
		System.out.println(temp);
		cell.setCellValue(content);
		in.close();
		FileOutputStream out = null;
        try {
            out = new FileOutputStream(temp);
            xwb.write(out);
            FileUtils.deleteFile(path);
             file = new File(temp);
             file.renameTo(new File(path));
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
		return true;
	}

	@Override
	public boolean writeCells(String path, Set<ExcelPosition> set) throws  Exception {
		InputStream in = new FileInputStream(path);
		XSSFWorkbook xwb = new XSSFWorkbook(in);
		XSSFSheet xSheet = xwb.getSheetAt(0);
		for(ExcelPosition ep : set){
			XSSFRow cellRow = xSheet.getRow(ep.getRow());
			if(cellRow ==null){
				cellRow = xSheet.createRow(ep.getRow());
			}
			XSSFCell cell = cellRow.getCell(ep.getCol());
			if(cell == null) {
				cell = cellRow.createCell(ep.getCol());
			}
			cell.setCellValue(ep.getContent());
		}
		
		FileOutputStream out = null;
        try {
            out = new FileOutputStream(path);
            xwb.write(out);
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
	public boolean writeCells(String srcpath, String desPath,
			Set<ExcelPosition> set) throws Exception {
		InputStream in = new FileInputStream(srcpath);
		XSSFWorkbook xwb = new XSSFWorkbook(in);
		XSSFSheet xSheet = xwb.getSheetAt(0);
		for(ExcelPosition ep : set){
			XSSFRow cellRow = xSheet.getRow(ep.getRow());
			if(cellRow ==null){
				cellRow = xSheet.createRow(ep.getRow());
			}
			XSSFCell cell = cellRow.getCell(ep.getCol());
			if(cell == null) {
				cell = cellRow.createCell(ep.getCol());
			}
			cell.setCellValue(ep.getContent());
		}
		
		FileOutputStream out = null;
        try {
            out = new FileOutputStream(desPath);
            xwb.write(out);
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
	public String toHtml(String path) {
		// TODO Auto-generated method stub
		return null;
	}

}
