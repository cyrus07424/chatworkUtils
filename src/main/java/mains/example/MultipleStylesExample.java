package mains.example;

import java.io.FileOutputStream;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * MultipleStylesExample.<br>
 * https://stackoverflow.com/questions/73069508/in-streaming-xssfworkbook-make-part-of-cell-content-to-bold-using-apache-poi
 *
 * @author cyrus
 */
public class MultipleStylesExample {
	public static void main(String[] args) {
		try (Workbook wb = new SXSSFWorkbook(new XSSFWorkbook(), 100, false, true)) {
			Sheet sheet = wb.createSheet("Sheet");
			Row row = sheet.createRow(2);
			XSSFFont font1 = (XSSFFont) wb.createFont();
			XSSFFont font2 = (XSSFFont) wb.createFont();
			XSSFFont font3 = (XSSFFont) wb.createFont();

			// cell1
			SXSSFCell hssfCell = (SXSSFCell) row.createCell(1);

			RichTextString richString = wb.getCreationHelper().createRichTextString("Hello, World!");
			richString.applyFont(0, 6, font1);
			richString.applyFont(6, 13, font2);
			hssfCell.setCellValue(richString);

			// cell2
			SXSSFCell cell = (SXSSFCell) row.createCell(2);
			XSSFRichTextString rt = new XSSFRichTextString("This is javatpoint");
			font1.setBold(true);
			font1.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
			rt.applyFont(0, 12, font1);
			font2.setItalic(true);
			font2.setUnderline(XSSFFont.U_DOUBLE);
			font2.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
			rt.applyFont(12, 18, font2);
			font3.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
			rt.append(" Learn New Technology Easily", font3);
			cell.setCellValue(rt);

			try (FileOutputStream os = new FileOutputStream("data/Javatpoint.xlsx")) {
				wb.write(os);
			}
		} catch (Exception e) {
			System.out.println(e);
		}
	}
}