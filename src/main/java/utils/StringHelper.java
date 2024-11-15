package utils;

/**
 * 文字列ヘルパー.
 *
 * @author cyrus
 */
public class StringHelper {

	/**
	 * 改行を表す正規表現.
	 */
	private static final String BREAK_CHARS_REGEX = "\r\n|\r|\n";

	/**
	 * 文字列を改行で分割.
	 *
	 * @param text
	 * @return
	 */
	public static String[] splitBreak(String text) {
		return text.split(BREAK_CHARS_REGEX);
	}
}