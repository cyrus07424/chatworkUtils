package mains.example;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * リッチテキストテスト.
 *
 * @author cyrus
 */
public class RichTextStringExample2 {
	public static void main(String[] args) throws FileNotFoundException, IOException {
		// 出力先ファイルを開く
		try (FileOutputStream fileOutputStream = new FileOutputStream(
				File.createTempFile("example", "RichTextStringExample2.xlsx", new File("data")))) {
			try (Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(), SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false,
					true)) {
				// シートを作成
				Sheet sheet = workbook.createSheet();
				((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();

				// フォントを作成
				Font headerFont = workbook.createFont();

				// セルスタイルを作成
				CellStyle headerRowCellStyle = workbook.createCellStyle();
				headerRowCellStyle.setFont(headerFont);

				// ヘッダー行1を作成
				List<Object> headerRowDataList = new ArrayList<>();
				headerRowDataList.add("あああ");

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
				headerRowDataList.add(richTextString1);
				createRow(sheet, headerRowCellStyle, headerRowDataList);

				// ヘッダー行2を作成
				List<Object> headerRowDataList2 = new ArrayList<>();
				headerRowDataList2.add("あああ");

				// 本文用のリッチテキストを作成2
				XSSFRichTextString richTextString2 = (XSSFRichTextString) workbook
						.getCreationHelper().createRichTextString("");
				System.out.println(richTextString2.getString());
				System.out.println(richTextString2.toString());
				richTextString2.append("> AAAあいうえお\n> ううう");
				System.out.println(richTextString2.getString());
				System.out.println(richTextString2.toString());
				richTextString2.append("BBB");
				System.out.println(richTextString2.getString());
				System.out.println(richTextString2.toString());
				headerRowDataList2.add(richTextString2);
				createRow(sheet, headerRowCellStyle, headerRowDataList2);

				// ヘッダー行2を作成
				List<Object> headerRowDataList3 = new ArrayList<>();
				headerRowDataList3.add("あああ");

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
				headerRowDataList3.add(richTextString3);
				createRow(sheet, headerRowCellStyle, headerRowDataList3);

				// 出力先ファイルに書き込み
				workbook.write(fileOutputStream);
			} catch (Exception e) {
				System.out.println(e);
			}
		}
	}

	/**
	 * シートに行を追加.
	 *
	 * @param sheet
	 * @param cellStyle
	 * @param dataList
	 */
	private static void createRow(Sheet sheet, CellStyle cellStyle, List<Object> dataList) {
		Row row = sheet.createRow(sheet.getLastRowNum() + 1);

		// 全てのデータに対して実行
		int columnIndex = 0;
		for (Object data : dataList) {
			// セルを作成
			SXSSFCell cell = (SXSSFCell) row.createCell(columnIndex++);
			cell.setCellStyle(cellStyle);

			// 値を設定
			if (data == null) {
				// NOP
			} else if (data instanceof String) {
				cell.setCellValue((String) data);
			} else if (data instanceof XSSFRichTextString) {
				cell.setCellValue((XSSFRichTextString) data);
			} else if (data instanceof RichTextString) {
				cell.setCellValue((RichTextString) data);
			} else if (data instanceof Integer) {
				cell.setCellValue((Integer) data);
			} else if (data instanceof Long) {
				cell.setCellValue((Long) data);
			} else if (data instanceof Float) {
				cell.setCellValue((Float) data);
			} else if (data instanceof Double) {
				cell.setCellValue((Double) data);
			} else if (data instanceof Date) {
				// FIXME
				cell.setCellValue(new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format((Date) data));
			} else {
				cell.setCellValue(String.valueOf(data));
			}
		}
	}
}