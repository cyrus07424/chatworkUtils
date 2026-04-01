package mains;

import java.io.File;
import java.io.FileOutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

import org.apache.commons.lang3.ArrayUtils;
import org.apache.commons.lang3.RegExUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.Strings;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FontUnderline;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.microsoft.playwright.Browser;
import com.microsoft.playwright.BrowserContext;
import com.microsoft.playwright.BrowserType;
import com.microsoft.playwright.Locator;
import com.microsoft.playwright.Page;
import com.microsoft.playwright.Playwright;
import com.microsoft.playwright.Response;
import com.microsoft.playwright.impl.TargetClosedError;
import com.microsoft.playwright.options.LoadState;

import constants.Configurations;
import utils.PlaywrightHelper;
import utils.StringHelper;

/**
 * チャットの全てのログをダウンロード.
 *
 * @author cyrus
 */
public class DownloadAllChatLogs {

	/**
	 * ダウンロードするチャットのID一覧.<br>
	 * (!rid000000000のような形式).
	 */
	private static final String[] TARGET_CHAT_ID_ARRAY = { "CHANGE ME" };

	/**
	 * デバッグモード.
	 */
	private static final boolean DEBUG_MODE = true;

	/**
	 * 最大ループ回数.
	 */
	private static final int MAX_LOOP_COUNT = Integer.MAX_VALUE;

	/**
	 * メイン.
	 *
	 * @param args
	 */
	public static void main(String[] args) {
		System.out.println("■start.");

		// Playwrightを作成
		try (Playwright playwright = Playwright.create()) {
			// ブラウザ起動オプションを取得
			BrowserType.LaunchOptions launchOptions = PlaywrightHelper.getLaunchOptions();

			// ブラウザを起動
			try (Browser browser = playwright.chromium().launch(launchOptions)) {
				System.out.println("■launched");

				// ブラウザコンテキストオプションを取得
				Browser.NewContextOptions newContextOptions = PlaywrightHelper.getNewContextOptions(true);

				// ブラウザコンテキストを取得
				try (BrowserContext context = browser.newContext(newContextOptions)) {
					// 全てのチャットのIDに対して実行
					for (String chatId : TARGET_CHAT_ID_ARRAY) {
						// ページを取得
						try (Page page = context.newPage()) {
							// チャット画面を表示
							Response response = page
									.navigate(String.format("https://www.chatwork.com/#%s-1", chatId));

							try {
								// 読み込み完了まで待機
								page.waitForLoadState(LoadState.NETWORKIDLE);
							} catch (Exception e) {
								// NOP
							}

							// 別スレッドで実行
							Thread thread = new Thread(() -> {
								int count = 0;
								while (true) {
									try {
										count++;

										// 一番上までスクロール
										page.evaluate("_timeLine.children[0].scrollTo(0, 0);");

										// 一番下までスクロール
										page.evaluate(
												"_timeLine.children[0].scrollTo(0, _timeLine.children[0].scrollHeight);");

										// 読み込み完了まで待機
										page.waitForLoadState(LoadState.NETWORKIDLE);

										// FIXME
										Thread.sleep(1000 + (250 * count));
									} catch (TargetClosedError e) {
										System.err.println("■target closed");
										break;
									} catch (InterruptedException e) {
										System.err.println("■sleep interrupted");
										break;
									} catch (Exception e) {
										System.err.println("■" + e.getMessage());
										e.printStackTrace();
									}
								}
							});
							thread.start();

							// 現在時刻を取得
							String dateString = new SimpleDateFormat("yyyyMMddHHmm").format(new Date());

							// チャット名
							String roomTitle = chatId;
							try {
								// チャット名を取得
								roomTitle = page.locator("#_roomTitle").textContent();
								roomTitle = RegExUtils.replaceAll(roomTitle, "[\\/]", "_");
								if (DEBUG_MODE) {
									System.out.println("■roomTitle: " + roomTitle);
								}
							} catch (Exception e) {
								e.printStackTrace();
							}

							// 出力先ファイルを作成
							File dataFile = new File(String.format("data/%s_%s.xlsx", roomTitle, dateString));

							// 出力先ファイルが存在する場合
							if (dataFile.exists()) {
								System.out.println("ファイルが既に存在します: " + dataFile.getAbsolutePath());

								// 終了
								break;
							}

							// 出力先ファイルを開く
							try (FileOutputStream fileOutputStream = new FileOutputStream(dataFile)) {
								// ブックを作成(共有文字列テーブルを使用)
								// https://stackoverflow.com/questions/73069508/in-streaming-xssfworkbook-make-part-of-cell-content-to-bold-using-apache-poi
								try (Workbook workbook = new SXSSFWorkbook(new XSSFWorkbook(),
										SXSSFWorkbook.DEFAULT_WINDOW_SIZE, false, true)) {
									try {
										// シートを作成
										Sheet sheet = workbook.createSheet();
										((SXSSFSheet) sheet).trackAllColumnsForAutoSizing();

										// フォントを作成
										Font headerFont = workbook.createFont();
										headerFont.setFontName(Configurations.BASE_FONT_NAME);
										headerFont.setBold(true);

										XSSFFont toAllFont = (XSSFFont) workbook.createFont();
										toAllFont.setFontName(Configurations.BASE_FONT_NAME);
										toAllFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
										toAllFont.setBold(true);

										XSSFFont toFont = (XSSFFont) workbook.createFont();
										toFont.setFontName(Configurations.BASE_FONT_NAME);
										toFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
										toFont.setBold(true);

										XSSFFont reFont = (XSSFFont) workbook.createFont();
										reFont.setFontName(Configurations.BASE_FONT_NAME);
										reFont.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
										reFont.setBold(true);

										XSSFFont quoteFont = (XSSFFont) workbook.createFont();
										quoteFont.setFontName(Configurations.BASE_FONT_NAME);
										quoteFont.setColor(HSSFColor.HSSFColorPredefined.GREY_50_PERCENT.getIndex());
										quoteFont.setItalic(true);
										quoteFont.setFontHeight(10);

										XSSFFont cwtagFont = (XSSFFont) workbook.createFont();
										cwtagFont.setFontName(Configurations.BASE_FONT_NAME);
										cwtagFont.setColor(HSSFColor.HSSFColorPredefined.BLUE.getIndex());
										cwtagFont.setBold(true);

										XSSFFont infoFont = (XSSFFont) workbook.createFont();
										infoFont.setFontName(Configurations.BASE_FONT_NAME);
										infoFont.setColor(HSSFColor.HSSFColorPredefined.GREEN.getIndex());
										infoFont.setBold(true);

										XSSFFont linkFont = (XSSFFont) workbook.createFont();
										linkFont.setFontName(Configurations.BASE_FONT_NAME);
										linkFont.setColor(HSSFColor.HSSFColorPredefined.LIGHT_BLUE.getIndex());
										linkFont.setUnderline(FontUnderline.SINGLE);

										XSSFFont codeFont = (XSSFFont) workbook.createFont();
										codeFont.setFontName(Configurations.BASE_FONT_NAME);
										codeFont.setColor(HSSFColor.HSSFColorPredefined.DARK_BLUE.getIndex());
										codeFont.setBold(true);
										codeFont.setFontHeight(10);

										XSSFFont notificationFont = (XSSFFont) workbook.createFont();
										notificationFont.setFontName(Configurations.BASE_FONT_NAME);
										notificationFont.setColor(HSSFColor.HSSFColorPredefined.BLUE_GREY.getIndex());
										notificationFont.setBold(true);

										// セルスタイルを作成
										CellStyle headerRowCellStyle = workbook.createCellStyle();
										headerRowCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
										headerRowCellStyle.setBorderBottom(BorderStyle.THIN);
										headerRowCellStyle.setBorderTop(BorderStyle.THIN);
										headerRowCellStyle.setBorderLeft(BorderStyle.THIN);
										headerRowCellStyle.setBorderRight(BorderStyle.THIN);
										headerRowCellStyle.setFillForegroundColor(
												HSSFColor.HSSFColorPredefined.LIGHT_GREEN.getIndex());
										headerRowCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
										headerRowCellStyle.setFont(headerFont);

										CellStyle dataRowCellStyle = workbook.createCellStyle();
										dataRowCellStyle.setVerticalAlignment(VerticalAlignment.TOP);
										dataRowCellStyle.setBorderBottom(BorderStyle.THIN);
										dataRowCellStyle.setBorderTop(BorderStyle.THIN);
										dataRowCellStyle.setBorderLeft(BorderStyle.THIN);
										dataRowCellStyle.setBorderRight(BorderStyle.THIN);
										dataRowCellStyle.setWrapText(true);

										// ヘッダー行を作成
										List<Object> headerRowDataList = new ArrayList<>();
										headerRowDataList.add("mid");
										headerRowDataList.add("index");
										headerRowDataList.add("投稿者");
										headerRowDataList.add("本文");
										headerRowDataList.add("投稿時間");
										headerRowDataList.add("deleted");
										headerRowDataList.add("bookmarked");
										createRow(sheet, headerRowCellStyle, headerRowDataList);

										// ウィンドウ枠の固定
										sheet.createFreezePane(0, 1);

										// ヘッダ行にオートフィルタの設定
										sheet.setAutoFilter(
												new CellRangeAddress(0, 0, 0, headerRowDataList.size() - 1));

										try {
											// 処理済みメッセージID一覧
											Set<Long> processedMidSet = new HashSet<>();

											// ループ内で処理済みの投稿数
											int processedMidCount;
											int loopCount = 0;
											while (true) {
												processedMidCount = 0;
												loopCount++;

												// 全てのメッセージを取得
												List<Locator> messageLocatorList = page
														.locator("#_mainContent #_timeLine [data-mid][data-index]")
														.all();

												// 全てのメッセージに対して実行
												for (Locator messageLocator : messageLocatorList) {
													// メッセージIDを取得
													Long mid = Long.parseLong(messageLocator.getAttribute("data-mid"));

													// 未処理メッセージの場合
													if (!processedMidSet.contains(mid)) {
														System.out.println("■mid: " + mid);

														// indexを取得
														Integer index = null;
														if (StringUtils.isNotBlank(
																messageLocator.getAttribute("data-index"))) {
															index = Integer.parseInt(
																	messageLocator.getAttribute("data-index"));
															if (DEBUG_MODE) {
																System.out.println("index: " + index);
															}
														}

														// deletedを取得
														Integer deleted = null;
														if (StringUtils.isNotBlank(
																messageLocator.getAttribute("data-deleted"))) {
															deleted = Integer.parseInt(
																	messageLocator.getAttribute("data-deleted"));
															if (DEBUG_MODE) {
																System.out.println("deleted: " + deleted);
															}
														}

														// bookmarkedを取得
														Integer bookmarked = null;
														if (StringUtils.isNotBlank(
																messageLocator.getAttribute("data-bookmarked"))) {
															bookmarked = Integer.parseInt(
																	messageLocator.getAttribute("data-bookmarked"));
															if (DEBUG_MODE) {
																System.out.println("bookmarked: " + bookmarked);
															}
														}

														// 投稿者を取得
														String userName = null;
														Locator userNameLocator = messageLocator
																.locator("[data-testid='timeline_user-name']");
														if (0 < userNameLocator.count()) {
															userName = userNameLocator.textContent();
															if (DEBUG_MODE) {
																System.out.println("userName: " + userName);
															}
														}

														// 本文用のリッチテキストを作成
														XSSFRichTextString textContentRichText = (XSSFRichTextString) workbook
																.getCreationHelper()
																.createRichTextString(StringUtils.EMPTY);
														// 削除済みでない場合
														if (deleted == null || deleted != 1) {
															// preタグのクラスを取得
															String[] preClassArray = messageLocator.locator("pre")
																	.first().getAttribute("class").split(" ");

															// preタグ直下の全ての要素に対して実行
															boolean infoAppended = false;
															for (Locator preInnerElementLocator : messageLocator
																	.locator("pre > *").all()) {
																// 要素のテキストを取得
																String innerText = StringHelper.normalizeNFKC(
																		preInnerElementLocator.innerText());

																// 要素の属性を取得
																String clazz = preInnerElementLocator
																		.getAttribute("class");
																String cwtag = preInnerElementLocator
																		.getAttribute("data-cwtag");
																String cwopen = preInnerElementLocator
																		.getAttribute("data-cwopen");
																String cwclose = preInnerElementLocator
																		.getAttribute("data-cwclose");

																// 要素の属性によって処理を実行
																if (StringUtils.isNotBlank(cwtag)) {
																	if (Strings.CS.equals(cwtag, "[toall]")) {
																		// [INFO]要素の挿入直後の場合
																		if (infoAppended) {
																			textContentRichText.append("\n");
																			infoAppended = false;
																		}
																		textContentRichText.append("[TO ALL]",
																				toAllFont);
																	} else if (Strings.CS.startsWith(cwtag,
																			"[To:")) {
																		// [INFO]要素の挿入直後の場合
																		if (infoAppended) {
																			textContentRichText.append("\n");
																			infoAppended = false;
																		}
																		textContentRichText.append("[TO]", toFont);
																	} else if (Strings.CS.startsWith(cwtag,
																			"[rp")) {
																		Pattern rpPattern = Pattern.compile(
																				"\\[rp aid=(\\d+) to=(\\d+)-(\\d+)\\]");
																		Matcher reMatcher = rpPattern
																				.matcher(cwtag);
																		if (reMatcher.find()) {
																			String rpToMid = reMatcher.group(3);
																			if (processedMidSet
																					.contains(
																							Long.parseLong(
																									rpToMid))) {
																				// [INFO]要素の挿入直後の場合
																				if (infoAppended) {
																					textContentRichText
																							.append("\n");
																					infoAppended = false;
																				}
																				textContentRichText.append(
																						String.format("[RE %s]",
																								rpToMid),
																						reFont);
																			} else {
																				// [INFO]要素の挿入直後の場合
																				if (infoAppended) {
																					textContentRichText
																							.append("\n");
																					infoAppended = false;
																				}
																				textContentRichText.append(
																						String.format("[RE %s-%s]",
																								reMatcher.group(2),
																								reMatcher.group(3)),
																						reFont);
																			}
																		} else {
																			// [INFO]要素の挿入直後の場合
																			if (infoAppended) {
																				textContentRichText.append("\n");
																				infoAppended = false;
																			}
																			textContentRichText.append("[RE]",
																					reFont);
																		}
																	} else if (Strings.CS.startsWith(cwtag,
																			"[preview")) {
																		// NOP
																	} else if (Strings.CS.startsWith(cwtag,
																			"http")) {
																		// [INFO]要素の挿入直後の場合
																		if (infoAppended) {
																			textContentRichText.append("\n");
																			infoAppended = false;
																		}
																		// FIXME リンクを取得
																		textContentRichText.append(cwtag, linkFont);
																	} else if (StringUtils.isNotBlank(innerText)) {
																		// [INFO]要素の挿入直後の場合
																		if (infoAppended) {
																			textContentRichText.append("\n");
																			infoAppended = false;
																		}

																		// preタグにgPYPSfまたはgmvZBzクラスが追加されている場合
																		if (ArrayUtils.containsAny(preClassArray,
																				"gPYPSf", "gmvZBz")) {
																			textContentRichText.append(innerText,
																					notificationFont);
																		} else {
																			textContentRichText.append(innerText);
																		}
																	} else if (StringUtils
																			.isNotBlank(preInnerElementLocator
																					.getAttribute("alt"))) {
																		// FIXME alt属性が存在する場合
																		textContentRichText
																				.append(preInnerElementLocator
																						.getAttribute("alt"));
																	} else {
																		// NOP
																	}
																} else if (Strings.CS.equals(cwopen, "[code]")) {
																	// [INFO]要素の挿入直後の場合
																	if (infoAppended) {
																		textContentRichText.append("\n");
																		infoAppended = false;
																	}

																	// 直前の文字が改行文字でない場合
																	String textContentRichTextString = textContentRichText
																			.getString();
																	if (StringUtils
																			.isNotBlank(textContentRichTextString)
																			&& textContentRichTextString.charAt(
																					textContentRichTextString
																							.length()
																							- 1) != '\n') {
																		textContentRichText.append("\n");
																	}

																	textContentRichText.append(innerText, codeFont);

																	// [INFO]要素の挿入扱いとする
																	infoAppended = true;
																} else if (Strings.CS.contains(clazz,
																		"chatQuote")) {
																	// [INFO]要素の挿入直後の場合
																	if (infoAppended) {
																		textContentRichText.append("\n");
																		infoAppended = false;
																	}

																	// 直前の文字が改行文字でない場合
																	String textContentRichTextString = textContentRichText
																			.getString();
																	if (StringUtils
																			.isNotBlank(textContentRichTextString)
																			&& textContentRichTextString.charAt(
																					textContentRichTextString
																							.length()
																							- 1) != '\n') {
																		textContentRichText.append("\n");
																	}

																	// 引用テキストを修正
																	innerText = Arrays
																			.stream(StringHelper
																					.splitBreak(innerText))
																			.map(line -> String.format("> %s\n",
																					line))
																			.collect(Collectors.joining());
																	textContentRichText.append(innerText,
																			quoteFont);
																} else {
																	// FIXME 「プレビュー」ボタンのテキストを削除
																	if (0 < preInnerElementLocator.locator("a")
																			.getByText("プレビュー",
																					new Locator.GetByTextOptions()
																							.setExact(true))
																			.count()) {
																		innerText = StringUtils.remove(innerText,
																				"プレビュー");
																	}

																	// FIXME ボタンのテキストを削除
																	for (Locator button : preInnerElementLocator
																			.locator("button").all()) {
																		if (DEBUG_MODE) {
																			System.out.println(
																					"■button: "
																							+ button.innerText());
																		}
																		if (StringUtils
																				.isNotBlank(button.innerText())) {
																			// 正規化したボタンのテキストを削除
																			innerText = StringUtils.remove(
																					innerText,
																					StringHelper.normalizeNFKC(
																							button.innerText()));
																		}
																	}

																	if (Strings.CS.equals(cwopen, "[info]")) {
																		// 直前の文字が改行文字でない場合
																		String textContentRichTextString = textContentRichText
																				.getString();
																		if (StringUtils.isNotBlank(
																				textContentRichTextString)
																				&& textContentRichTextString.charAt(
																						textContentRichTextString
																								.length()
																								- 1) != '\n') {
																			textContentRichText.append("\n");
																		}

																		textContentRichText.append(innerText,
																				infoFont);
																		infoAppended = true;
																	} else {
																		// [INFO]要素の挿入直後の場合
																		if (infoAppended) {
																			textContentRichText.append("\n");
																			infoAppended = false;
																		}

																		// preタグにgPYPSfまたはgmvZBzクラスが追加されている場合
																		if (ArrayUtils.containsAny(preClassArray,
																				"gPYPSf", "gmvZBz")) {
																			textContentRichText.append(innerText,
																					notificationFont);
																		} else {
																			textContentRichText.append(innerText);
																		}
																	}
																}
															}

															if (DEBUG_MODE) {
																System.out.println("textContent: "
																		+ textContentRichText.getString());
															}
														} else {
															// FIXME
															Locator spanLocator = messageLocator
																	.locator("span[data-cwtag]");
															textContentRichText.append(spanLocator.textContent(),
																	cwtagFont);
															if (DEBUG_MODE) {
																System.out.println("■span: "
																		+ textContentRichText.getString());
															}
														}

														// 投稿時間を取得
														Date tmDate = null;
														Locator tmLocator = messageLocator.locator("[data-tm]");
														if (0 < tmLocator.count()) {
															Long tm = Long
																	.parseLong(tmLocator.getAttribute("data-tm"));
															tmDate = new Date(tm * 1000);
															if (DEBUG_MODE) {
																System.out.println("tm: " + tm);
																System.out.println("tmDate: " + tmDate);
															}
														}

														// 行データを作成
														List<Object> dataRowDataList = new ArrayList<>();
														dataRowDataList.add(String.format("%d", mid));
														dataRowDataList.add(index);
														dataRowDataList.add(userName);
														dataRowDataList.add(textContentRichText);
														dataRowDataList.add(tmDate);
														dataRowDataList.add(deleted);
														dataRowDataList.add(bookmarked);

														// 行を作成
														createRow(sheet, dataRowCellStyle, dataRowDataList);

														// 処理済みメッセージID一覧に追加
														processedMidSet.add(mid);
														processedMidCount++;
													} else {
														if (DEBUG_MODE) {
															System.out.println("■mid: " + mid + "(処理済み)");
														}
													}
												}

												if (MAX_LOOP_COUNT <= loopCount || processedMidCount == 0) {
													break;
												}
											}
										} finally {
											// 列幅の自動調整
											for (int i = 0; i <= headerRowDataList.size(); i++) {
												sheet.autoSizeColumn(i);
											}
										}
									} finally {
										// 出力先ファイルに書き込み
										workbook.write(fileOutputStream);
									}
								}
							} finally {
								// 別スレッドを停止
								thread.interrupt();
							}
						}
					}
				}
			}
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			System.out.println("■done.");
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