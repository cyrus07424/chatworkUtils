package constants;

import java.io.File;
import java.nio.file.Path;

/**
 * 環境設定.
 *
 * @author cyrus
 */
public interface Configurations {

	/**
	 * ブラウザをヘッドレスモードで使用するか.
	 */
	boolean USE_HEADLESS_MODE = false;

	/**
	 * ブラウザのステートの出力先.
	 */
	Path STATE_PATH = new File("data/state.json").toPath();

	/**
	 * 基本のフォント名.
	 */
	String BASE_FONT_NAME = "游ゴシック";

	/**
	 * 出力先ファイルを上書きするかどうか.
	 */
	boolean OVERWRITE_DATA_FILE = true;
}