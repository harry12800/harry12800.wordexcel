package com.newtec.tree2word.word;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordTool {

	/**
	 * 文件路径filePath该文件是否是word文件 包含 word2003 和word2007
	 * 
	 * @param filePath
	 *            文件路径
	 * @return
	 */
	public static boolean isWordFile(String filePath) {
		int word2003 = 0;
		int word2007 = 0;
		/* 2003 */
		try (FileInputStream f = new FileInputStream(filePath)){
			new HWPFDocument(f);
		} catch (FileNotFoundException e) {
			return false;
		} catch (Exception e) {
			word2003 = 1;
		}
		/* 2007 */
		try (FileInputStream f = new FileInputStream(filePath)){
			new XWPFDocument(f);
		} catch (FileNotFoundException e) {
			return false;
		} catch (Exception e) {
			word2007 = 1;
		}
		if (word2003 + word2007 == 1)
			return true;
		return false;
	}
	/**
	 * 文件路径filePath 包含 word2003 和word2007 两个版本
	 * @param filePath 文件路径
	 * @return 1 代表是2003 2代表是 2007 ，其他代表两者都不是
	 */
	public static int getWordVersion(String filePath) {
		int word2003 = 0;
		int word2007 = 0;
		/* 2003 */
		try (FileInputStream f = new FileInputStream(filePath)){
			new HWPFDocument(f);
		} catch (FileNotFoundException e) {
			return -1;
		} catch (Exception e) {
			word2003 = 1;
		}
		/* 2007 */
		try (FileInputStream f = new FileInputStream(filePath)){
			new XWPFDocument(f);
		} catch (FileNotFoundException e) {
			return -1;
		} catch (Exception e) {
			word2007 = 1;
		}
		if (word2003 == 0)
			return 1;
		if (word2007 == 0)
			return 2;
		return -1;
	}
	
	public static void main(String[] args) {
		if(WordTool.isWordFile("C:\\Users\\harry12800\\Desktop\\0407 火焰光度法.docx")){
			System.out.println("是");
		}
		else{
			System.out.println("否");
		}
		
		if(1==WordTool.getWordVersion("C:\\Users\\harry12800\\Desktop\\0407 火焰光度法.docx")){
			System.out.println("2003");
		}
		else if(2==WordTool.getWordVersion("C:\\Users\\harry12800\\Desktop\\0407 火焰光度法.docx")){
			System.out.println("2007");
		}
		else{
			System.out.println("都不是");
		}
	}
	public static boolean isWordFile(File zipFile) {
		return isWordFile(zipFile.getAbsolutePath());
	}
}
