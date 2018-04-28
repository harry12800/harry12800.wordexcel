package com.newtec.tree2word.word;

 

public class Word2007 implements Word2html {

	@Override
	public String tohtml(String path,String tmpFileDir) {
		
//		try (ByteOutputStream out = new ByteOutputStream();
//				InputStream in = new FileInputStream(path);){
//			
//			XWPFDocument document = new XWPFDocument(in);
//			//List<IBodyElement> list = document.getBodyElements();
////			for (IBodyElement iBodyElement : list) {
////				//XWPFParagraph paragraph = (XWPFParagraph) iBodyElement;
////				//paragraph.setStyle("space");
////			}
//			// 转换后html中图片src的链接
//			XHTMLOptions options = XHTMLOptions.create().URIResolver(new BasicURIResolver(tmpFileDir));
//			File imageFolderFile = new File(tmpFileDir);
//			options.setExtractor(new FileImageExtractor(imageFolderFile));
//			// 3) Convert XWPFDocument to XHTML
//			XHTMLConverter.getInstance().convert(document, out, options);
//			System.out.println(new String(out.getBytes()));
//			return new String(out.getBytes());
//		} catch (Exception e) {
//			e.printStackTrace();
//		}
		return "";
	}

	public static void main(String[] args) {
		String html = new Word2007().tohtml("C:\\Users\\Administrator\\Desktop\\新建 Microsoft Office Word 2007 文档.docx","D://");
		System.out.println(html);
	}
}
