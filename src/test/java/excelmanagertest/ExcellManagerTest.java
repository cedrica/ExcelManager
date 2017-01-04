package excelmanagertest;

import static org.junit.Assert.*;

import java.util.ArrayList;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.junit.Before;
import org.junit.Test;

import excelmanager.ExcellManager;


public class ExcellManagerTest {

	private ExcellManager excelmanager;
	private List<TestObject> testObjects = null;
	@Before
	public void init(){
		excelmanager = new ExcellManager();
		testObjects =  createSomeTestObjects(10);
	}
	
	private List<TestObject> createSomeTestObjects(int count) {
		List<TestObject> testObjects = new ArrayList<TestObject>();
		for (int id = 0; id < count; id++) {
			TestObject testObject = new TestObject(id, "vorname"+id, "nachname"+id, id*10, "addresse"+id);
			testObjects.add(testObject);
		}
		return testObjects;
	}
	
	@Test
	public void testGenerateSingleReportSheet() {
		HSSFWorkbook wb = excelmanager.generateSingleReportSheet(null);
		assertNull(wb);
		wb = excelmanager.generateSingleReportSheet(testObjects);
		assertEquals(wb.getSheetAt(0).getSheetName(), "test");
	}


}
