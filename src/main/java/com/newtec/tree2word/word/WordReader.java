package com.newtec.tree2word.word;

import java.io.FileInputStream;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

public class WordReader {
	private static Word2003 word2003 = new Word2003();
	private static Word2007 word2007 = new Word2007();
	
	
	public static String toHtml(String path,String tmpFileDir){
		if(isWord2003(path)){
			return word2003.tohtml(path,tmpFileDir);
		}else if(isWord2007(path)){
			return word2007.tohtml(path,tmpFileDir);
		}
		return "文件错误！";
	}

	public static void main(String[] args) {
		String str = WordReader.toHtml("C:\\Users\\Yuexin\\Desktop\\数据压缩（二）.docx","D:\\infoshar1");
		System.out.println(str);
	}

	private static boolean isWord2003(String path) {
		try (FileInputStream file=new FileInputStream(path);){
			new HWPFDocument(file);
			return true;
		} catch (Exception e) {
			return false;
		}
	}


	private static boolean isWord2007(String path) {
		 try (FileInputStream file=new FileInputStream(path);){
			new XWPFDocument(file);
			return true;
		} catch (Exception e) {
			return false;
		}
	}
}
