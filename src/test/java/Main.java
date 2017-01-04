import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import excelmanager.ExcellManager;
import javafx.application.Application;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

public class Main extends Application{

	public Main() {

	}

	private List<TestObject> createSomeTestObjects(int count) {
		List<TestObject> testObjects = new ArrayList<TestObject>();
		for (int id = 0; id < count; id++) {
			TestObject testObject = new TestObject(id, "vorname"+id, "nachname"+id, id*10, "addresse"+id);
			testObjects.add(testObject);
		}
		return testObjects;
	}

	public static void main(String[] args) {
		launch(args);
	}

	@Override
	public void start(Stage primaryStage) throws Exception {
		List<TestObject> testObjects =  createSomeTestObjects(10);
		ExcellManager excellManager = new ExcellManager();
		HSSFWorkbook workbook = excellManager.generateSingleReportSheet(testObjects);
		FileChooser fileChooser = new FileChooser();
		FileChooser.ExtensionFilter extFilter = new FileChooser.ExtensionFilter("XLS files (*.xls)", "*.xls");
        fileChooser.getExtensionFilters().add(extFilter);
		File fileToBeSaved = fileChooser.showSaveDialog(primaryStage);
		
		if (fileToBeSaved != null) {
			try {
				File f = new File(fileToBeSaved.getPath());
				if(f.isFile() && f.exists()){
					f.delete();
				}
				FileOutputStream out = new FileOutputStream(f,false);
				workbook.write(out);
				out.close();
				System.out.println("Excel written successfully..");
			} catch (IOException e) {
				e.printStackTrace();
			}
			System.out.println("File saved: " + fileToBeSaved.getPath());
		} else {
			System.err.println("ERROR: File path is null.");
		}
	}
}
