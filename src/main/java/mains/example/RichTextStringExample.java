package mains.example;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * リッチテキストテスト.
 *
 * @author cyrus
 */
public class RichTextStringExample {
	public static void main(String[] args) {
		try (Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false,
				true)) {

			// 本文用のリッチテキストを作成1
			XSSFRichTextString richTextString1 = (XSSFRichTextString) workbook
					.getCreationHelper().createRichTextString(null);
			System.out.println(richTextString1.getString());
			System.out.println(richTextString1.toString());
			richTextString1.append("AAA");
			System.out.println(richTextString1.getString());
			System.out.println(richTextString1.toString());
			richTextString1.append("BBB");
			System.out.println(richTextString1.getString());
			System.out.println(richTextString1.toString());

			// 本文用のリッチテキストを作成2
			XSSFRichTextString richTextString2 = (XSSFRichTextString) workbook
					.getCreationHelper().createRichTextString("");
			System.out.println(richTextString2.getString());
			System.out.println(richTextString2.toString());
			richTextString2.append("AAA");
			System.out.println(richTextString2.getString());
			System.out.println(richTextString2.toString());
			richTextString2.append("BBB");
			System.out.println(richTextString2.getString());
			System.out.println(richTextString2.toString());

			// 本文用のリッチテキストを作成3
			XSSFRichTextString richTextString3 = (XSSFRichTextString) workbook
					.getCreationHelper().createRichTextString("333");
			System.out.println(richTextString3.getString());
			System.out.println(richTextString3.toString());
			richTextString3.append("AAA");
			System.out.println(richTextString3.getString());
			System.out.println(richTextString3.toString());
			richTextString3.append("BBB");
			System.out.println(richTextString3.getString());
			System.out.println(richTextString3.toString());
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}