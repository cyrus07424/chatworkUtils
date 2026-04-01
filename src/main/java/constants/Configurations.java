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
	 * 使用するユーザーエージェント.
	 */
	String USE_USER_AGENT = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/133.0.0.0 Safari/537.36";

	/**
	 * ブラウザのステートの出力先.
	 */
	Path STATE_PATH = new File("data/state.json").toPath();

	/**
	 * 基本のフォント名.
	 */
	String BASE_FONT_NAME = "游ゴシック";
}