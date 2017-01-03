package excelmanager.exception;

import javafx.scene.control.Alert;
import javafx.scene.control.Alert.AlertType;

public class Assertion {

	public static boolean NDEBUG = true;

	private static void printStack(String why) {
		Throwable t = new Throwable(why);
		t.printStackTrace();

		System.exit(1);
	}

	public static void _assert(boolean expression, String why) {
		if (NDEBUG && !expression) {
			Alert alert = new Alert(AlertType.ERROR);
			alert.setTitle("Error Dialog");
			alert.setHeaderText("Look, an Error Dialog");
			alert.setContentText(why);
			alert.showAndWait();
			printStack(why);
		}
	}
}
