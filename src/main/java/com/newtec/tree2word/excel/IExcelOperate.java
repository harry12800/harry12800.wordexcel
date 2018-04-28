package com.newtec.tree2word.excel;

import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.Set;

public interface IExcelOperate {
	/**
	 * 读取excel第一行
	 * @param path
	 * @return
	 */
	public String[] readExcelTitle(String path);
	/**
	 * 读取excel第二行和之后的数据
	 * @param path
	 * @return
	 */
	public Map<Integer, List<String>> readExcelContent(String path);
	/**
	 * 修改excel的第row行第col列的内容
	 * @param path
	 * @param row
	 * @param col
	 * @param content
	 * @return
	 * @throws IOException
	 */
	public boolean writeCell(String path, Integer row, int col,String content) throws IOException;
	/**
	 * 删除行
	 * @param path
	 * @param row
	 * @return
	 */
	public boolean deleteRow(String path,int row);
	/**
	 * 批量修改单元格内容
	 * @param path
	 * @param set
	 * @return
	 * @throws Exception
	 */
	public boolean writeCells(String path,Set<ExcelPosition> set) throws Exception;
	public boolean writeCells(String srcpath,String desPath,Set<ExcelPosition> set) throws Exception;
	public String toHtml(String path);
}
