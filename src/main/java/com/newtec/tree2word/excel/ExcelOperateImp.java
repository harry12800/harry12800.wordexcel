package com.newtec.tree2word.excel;

import java.io.IOException;
import java.util.List;
import java.util.Map;

public interface ExcelOperateImp {
	public String[] readExcelTitle(String path);
	public Map<Integer, List<String>> readExcelContent(String path);
	public Object writeCell(String path, Integer row, int changeIndex,String content) throws IOException;
}
