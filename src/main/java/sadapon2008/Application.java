package sadapon2008;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.nio.file.StandardCopyOption;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xssf.usermodel.XSSFDrawing;
import org.apache.poi.xssf.usermodel.XSSFShape;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFSimpleShape;
import org.apache.poi.xssf.usermodel.XSSFTextParagraph;
import org.apache.poi.xssf.usermodel.XSSFTextRun;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;

public class Application {

	public static void main(String[] args) {
		// コマンドライン引数のチェックと取得
		if (args.length < 4) {
			System.exit(1);
		}
		// 置換対象の文字列
		String textTarget = args[0];
		// 置換後の文字列
		String textReplacement = args[1];
		// 置換対象の.xlsxファイル
		String filenameSrc = args[2];
		// 置換後に作成される.xlsxファイル
		String filenameDest = args[3];
		
		try {
			// 出力先ファイルをコピーしてから書き換える
			Files.copy(Paths.get(filenameSrc), Paths.get(filenameDest), StandardCopyOption.REPLACE_EXISTING);
			OPCPackage pkg = OPCPackage.open(new FileInputStream(filenameDest));
			XSSFWorkbook workBook = new XSSFWorkbook(pkg);
			// シートごとに処理する
			int n = workBook.getNumberOfSheets();
			for (int i = 0; i < n; i++) {
				XSSFSheet sheet = workBook.getSheetAt(i);
				XSSFDrawing drawing = sheet.createDrawingPatriarch();
				// オートシェイプごとに処理する
				for (XSSFShape shape : drawing.getShapes()) {
					if (!(shape instanceof XSSFSimpleShape)) {
						// グループ化したオートシェイプには非対応
						continue;
					}
					// グループ化してないオートシェイプに対して処理する
					XSSFSimpleShape simpleShape = (XSSFSimpleShape)shape;
					CTTextBody textBody = simpleShape.getCTShape().getTxBody();
					if (null == textBody) {
						continue;
					}
					for (XSSFTextParagraph textParagraph : simpleShape.getTextParagraphs()) {
						for (XSSFTextRun textRun : textParagraph.getTextRuns()) {
							// 書式設定がされている単位のテキストを置換していく
							// そのため文字単位で書式設定が違う場合には非対応
							textRun.setText(textRun.getText().replace(textTarget, textReplacement));
						}
					}
				}
			}
			FileOutputStream fileOut = new FileOutputStream(filenameDest);
			workBook.write(fileOut);
			fileOut.close();
			workBook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}
}
