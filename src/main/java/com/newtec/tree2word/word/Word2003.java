package com.newtec.tree2word.word;

import java.io.ByteArrayOutputStream;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.net.URLDecoder;
import java.util.List;

import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.transform.OutputKeys;
import javax.xml.transform.Transformer;
import javax.xml.transform.TransformerFactory;
import javax.xml.transform.dom.DOMSource;
import javax.xml.transform.stream.StreamResult;

import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.converter.PicturesManager;
import org.apache.poi.hwpf.converter.WordToHtmlConverter;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.hwpf.usermodel.PictureType;
import org.w3c.dom.Document;

public class Word2003 implements Word2html {

	@Override
	public String tohtml(String path, String tmpDir) {
		try {
			return word2003Html(path, tmpDir, tmpDir);
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		return path;
	}

	public String word2003Html(String wordFileName, final String imgUrl,
			String picPath) throws Exception {
		try (FileInputStream fileInputStream = new FileInputStream(wordFileName);
				ByteArrayOutputStream out = new ByteArrayOutputStream();) {
			HWPFDocument wordDocument = new HWPFDocument(fileInputStream);
			WordToHtmlConverter wordToHtmlConverter = new WordToHtmlConverter(
					DocumentBuilderFactory.newInstance().newDocumentBuilder()
							.newDocument());
			wordToHtmlConverter
					.setPicturesManager(new MyPicturesManager(imgUrl));
			wordToHtmlConverter.processDocument(wordDocument);
			// save pictures
			List<?> pics = wordDocument.getPicturesTable().getAllPictures();
			if (pics != null) {
				for (int i = 0; i < pics.size(); i++) {
					Picture pic = (Picture) pics.get(i);
					//PictureType type = pic.suggestPictureType();
					try {
						System.out.println("---filename----" + picPath
								+ "---size:" + pic.getContent().length);
						pic.writeImageContent(new FileOutputStream("D:\\temp\\"
								+ pic.suggestFullFileName()));
						// if (type == PictureType.WMF) {
						// Wmf2Jpg.convert(ParamData.jghFilePath + picPath
						// + pic.suggestFullFileName());
						// }
					} catch (FileNotFoundException e) {
						e.printStackTrace();
					}
				}
			}
			Document htmlDocument = wordToHtmlConverter.getDocument();

			DOMSource domSource = new DOMSource(htmlDocument);
			StreamResult streamResult = new StreamResult(out);

			TransformerFactory tf = TransformerFactory.newInstance();
			Transformer serializer = tf.newTransformer();
			serializer.setOutputProperty(OutputKeys.ENCODING, "utf-8");
			// serializer.setOutputProperty(OutputKeys.ENCODING, "GB2312");
			serializer.setOutputProperty(OutputKeys.INDENT, "yes");
			serializer.setOutputProperty(OutputKeys.METHOD, "html");
			serializer.setOutputProperty(OutputKeys.DOCTYPE_PUBLIC, "String");
			serializer.transform(domSource, streamResult);
			String str = new String(out.toByteArray());
			str = str
					.replaceAll("<img src=(\"[^<]*)?.wmf\"",
							"<img src=$1.jpg\"")
					// 更换图片（公式的图片替换）
					.replaceAll("<p ", "<p style='white-space: pre-wrap;' ")
					// 使p标签的内容保持原来的样式
					// 如很多空格的情况
					.replaceAll("<body class=\"(.*?)\">",
							"<body style='margin:10px'>")
					// 去掉body上的样式
					.replaceAll("<div([^>]*?)>", "<p$1>")
					.replaceAll("</div>", "</p>");// 去掉div标签
			// System.err.println(str);
			return str;
			// writeFile(str, htmlFileName);
			// return getStyle2003(str);
		} catch (IllegalArgumentException e) {
			// return realRtfDoc(wordFileName, htmlFileName);
		} catch (Exception e) {
		}
		return "";
	}

	// private String realRtfDoc(String wordFileName, String htmlFileName)
	// throws CustomException {
	// String html = "";
	// try {
	// html = RTF2HTMLUtil.rtf2html(wordFileName);
	// html = html.replaceAll("<p ", "<p style='white-space: pre-wrap;' ");
	// html = html.replaceAll("<body class=\"b1 b2\">",
	// "<body style='margin:10px'>");
	// } catch (Exception e) {
	// throw new CustomException("", e.getMessage());
	// }
	// writeRtfHtmlFile(html, htmlFileName);
	// return "";
	// }
	class MyPicturesManager implements PicturesManager {
		private String imgUrl = "";

		public MyPicturesManager(String imgurl) {
			this.imgUrl = imgurl;
		}

		@SuppressWarnings("deprecation")
		@Override
		public String savePicture(byte[] content, PictureType pictureType,
				String suggestedName, float widthInches, float heightInches) {
			/**
			 * if判断wordToHtmlConverter读到的这张图片的字节长度如果为0 直接返回null
			 * 那么转换后html文件将当这个图片不存在
			 **/
			if (content.length == 0)
				return null;
			return URLDecoder.decode(imgUrl + suggestedName);
		}

	}
}
