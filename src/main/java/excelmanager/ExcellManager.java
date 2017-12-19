/**
 * this class is used to generate different kind of Excels report for PoJos
 * 
 * @author ca.leumaleu
 */
package excelmanager;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;

import excelmanager.annotations.XLS;
import excelmanager.annotations.XlsAdditionalInformation;
import excelmanager.annotations.XlsColumn;
import excelmanager.annotations.XlsStyler;
import excelmanager.enums.Location;
import excelmanager.enums.Orientation;
import excelmanager.style.StyleData;
import excelmanager.style.WoorkbookStyler;
import javafx.util.converter.CurrencyStringConverter;

public class ExcellManager {

	private HashMap<String, HSSFCellStyle> hmFieldStyle = null;
	private static int FRIST_ROW_INDEX_FOR_DATA = 0;

	public ExcellManager() {
	}

	/**
	 * generate a workbook for all POJOs inside the given parameter list.
	 *
	 * @param entities
	 */
	public <T> HSSFWorkbook generateSingleReportSheet(List<T> entities) {
		if (entities == null || entities.size() == 0)
			return null;
		FRIST_ROW_INDEX_FOR_DATA = 0;
		HSSFWorkbook workbook = new HSSFWorkbook();
		String sheetname = "";
		int rownum = 0;
		int cellnum = 0;
		HashMap<Location, String> footerInfo = new HashMap<>();
		Row row;
		Class<? extends Object> entityClazz = entities.get(0).getClass();
		XLS xls = (XLS) entityClazz.getAnnotation(XLS.class);
		XlsAdditionalInformation additionalInformation = null;
		if (xls != null) {
			sheetname = (xls.sheetsname().trim().length() <= 0) ? entityClazz.getSimpleName() : xls.sheetsname();
			additionalInformation = xls.xlsAdditionalInformation();
		} else {
			System.err.println("ERROR: Die Klasse " + entityClazz.getSimpleName()
					+ " wurde nicht als Excel-Report deklariert. Bitte die mit @XLS annotieren");
			return null;
		}
		HSSFSheet sheet = workbook.createSheet(sheetname);
		if (additionalInformation != null) {
			String text = additionalInformation.text();
			Location location = additionalInformation.location();
			int colspan = additionalInformation.colspan();
			if (location == Location.BOTTOM) {
				footerInfo.put(Location.BOTTOM, text + " colspan= " + colspan);
			} else if (!text.isEmpty()) {
				row = sheet.createRow(rownum++);
				FRIST_ROW_INDEX_FOR_DATA = rownum;
				Cell cell = row.createCell((short) 0);
				cell.setCellValue(text);
				mergeCells(sheet, colspan, 0);
			}
		}
		final List<Field> fields = Arrays.asList(entityClazz.getDeclaredFields());
		final List<Field> usedFields = new ArrayList<Field>();
		final List<String> nameOfUsedFields = new ArrayList<String>();
		hmFieldStyle = new HashMap<String, HSSFCellStyle>();
		WoorkbookStyler woorkbookStyler = new WoorkbookStyler();

		fields.forEach(f -> {
			XlsColumn xlsAnnotation = f.getAnnotation(XlsColumn.class);
			if (xlsAnnotation != null) {
				String customname = xlsAnnotation.customname();
				XlsStyler xlsStyler = xlsAnnotation.styler();
				StyleData styleData = extractStyleInfo(xlsStyler);
				woorkbookStyler.setStyle(styleData);
				woorkbookStyler.style(workbook);
				nameOfUsedFields.add((customname.length() <= 0) ? f.getName() : customname);
				usedFields.add(f);
				hmFieldStyle.put((customname.length() <= 0) ? f.getName() : customname, WoorkbookStyler.style);
			}
		});
		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
		data.put(0, nameOfUsedFields.toArray());
		int key = 1;
		for (T t : entities) {
			data.put(key++, builtRowFrom(t, usedFields));
		}

		insertDataIntoSheet(rownum, cellnum, sheet, data);
		if (footerInfo != null && footerInfo.size() > 0) {
			String textColspan = footerInfo.get(Location.BOTTOM);
			if (textColspan.trim().length() > 0) {
				String[] splittedStr = textColspan.split("colspan=");
				String text = splittedStr[0];
				int colspan = Integer.valueOf(splittedStr[1].trim());
				row = sheet.createRow(rownum++);
				Cell cell = row.createCell((short) 0);
				cell.setCellValue(text);
				int rowIndex = entities.size() + 1;
				mergeCells(sheet, colspan, rowIndex);
			}
		}
		return workbook;
	}

	/**
	 * generate a workbook for all POJOs inside the given parameter list.
	 *
	 * @param entities
	 */
	public <T> String generateSingleReportAsCSV(List<T> entities) {
		if (entities == null || entities.size() == 0)
			return null;
		FRIST_ROW_INDEX_FOR_DATA = 0;
		int rownum = 0;
		HashMap<Location, String> footerInfo = new HashMap<>();
		Class<? extends Object> entityClazz = entities.get(0).getClass();
		XLS xls = (XLS) entityClazz.getAnnotation(XLS.class);
		XlsAdditionalInformation additionalInformation = null;
		if (xls != null) {
//			String sheetname = (xls.sheetsname().trim().length() <= 0) ? entityClazz.getSimpleName() : xls.sheetsname();
			additionalInformation = xls.xlsAdditionalInformation();
		} else {
			System.err.println("ERROR: Die Klasse " + entityClazz.getSimpleName()
					+ " wurde nicht als Excel-Report deklariert. Bitte die mit @XLS annotieren");
			return null;
		}
		if (additionalInformation != null) {
			String text = additionalInformation.text();
			Location location = additionalInformation.location();
			int colspan = additionalInformation.colspan();
			if (location == Location.BOTTOM) {
				footerInfo.put(Location.BOTTOM, text + " colspan= " + colspan);
			} else if (!text.isEmpty()) {
				FRIST_ROW_INDEX_FOR_DATA = rownum;
			}
		}
		final List<Field> fields = Arrays.asList(entityClazz.getDeclaredFields());
		final List<Field> usedFields = new ArrayList<Field>();
		final List<String> nameOfUsedFields = new ArrayList<String>();

		fields.forEach(f -> {
			XlsColumn xlsAnnotation = f.getAnnotation(XlsColumn.class);
			if (xlsAnnotation != null) {
				String customname = xlsAnnotation.customname();
				nameOfUsedFields.add((customname.length() <= 0) ? f.getName() : customname);
				usedFields.add(f);
			}
		});
		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
		data.put(0, nameOfUsedFields.toArray());
		int key = 1;
		for (T t : entities) {
			data.put(key++, builtRowFrom(t, usedFields));
		}

		String result = insertDataIntoCSV(rownum, data);
		return result;
	}

	public void mergeCells(HSSFSheet sheet, int colspan, int rowIndex) {
		sheet.addMergedRegion(new CellRangeAddress(rowIndex, // mention first
																// row here
				rowIndex, // mention last row here, it is 1 as we are doing a
				0, // mention first column of merging
				colspan // mention last column to include in merge
		));
	}

	public String insertDataIntoCSV(int rownum, Map<Integer, Object[]> data) {
		String row = "";
		for (Map.Entry<Integer, Object[]> rowSet : data.entrySet()) {
			if (rownum == FRIST_ROW_INDEX_FOR_DATA) {
				for (Object obj : rowSet.getValue()) {
					row += obj + ";";
				}
				rownum++;
				row += "\n";
				continue;
			}
			for (Object obj : rowSet.getValue()) {
				if (obj == null){
					row += ";";
				}else if (obj instanceof Double){
					CurrencyStringConverter c = new CurrencyStringConverter();
					String item = c.toString(new BigDecimal(obj.toString()));
					row += item+";";
				}else
					row += obj.toString() + ";";
			}
			row += "\n";
			rownum++;
		}
		return row;
	}

	public void insertDataIntoSheet(int rownum, int cellnum, HSSFSheet sheet, Map<Integer, Object[]> data) {
		Row row;
		for (Map.Entry<Integer, Object[]> rowSet : data.entrySet()) {
			row = sheet.createRow(rownum);
			if (rownum == FRIST_ROW_INDEX_FOR_DATA) {
				for (Object obj : rowSet.getValue()) {
					Cell cell = row.createCell(cellnum++);
					cell.setCellValue(obj.toString());
					cell.setCellStyle(hmFieldStyle.get(obj.toString()));
				}
				rownum++;
				continue;
			}
			cellnum = 0;

			for (Object obj : rowSet.getValue()) {
				insertObjectAtRow(sheet, row, obj, cellnum);
				cellnum++;
			}
			rownum++;
		}
	}

	private StyleData extractStyleInfo(XlsStyler xlsStyler) {
		StyleData styleData = new StyleData();
		short bgColor = xlsStyler.bgColor();
		short fgColor = xlsStyler.fgColor();
		Orientation orientation = xlsStyler.orientation();
		boolean italic = xlsStyler.isItalic();
		boolean bold = xlsStyler.isBold();
		String fontName = xlsStyler.fontName();
		int fontSize = xlsStyler.fontSize();

		styleData.setOrientation(orientation);
		styleData.setBgColor(bgColor);
		styleData.setFgColor(fgColor);
		styleData.setBold(bold);
		styleData.setFontSize((short) fontSize);
		styleData.setItalic(italic);
		styleData.setFontName(fontName);
		return styleData;
	}

	private void insertObjectAtRow(HSSFSheet sheet, Row row, Object obj, int cellnum) {
		Cell cell = row.createCell(cellnum);
		if (obj instanceof Date) {
			cell.setCellValue((Date) obj);
		} else if (obj instanceof Boolean) {
			cell.setCellValue((Boolean) obj);
		} else if (obj instanceof String) {
			cell.setCellValue((String) obj);
		} else if (obj instanceof Double) {
			cell.setCellValue((Double) obj);
		} else if (obj instanceof Integer) {
			cell.setCellValue((Integer) obj);
		} else if (obj instanceof Long) {
			cell.setCellValue((Long) obj);
		} else if (obj instanceof Float) {
			cell.setCellValue((Float) obj);
		}
		sheet.autoSizeColumn((short) cellnum);
	}

	private <T> Object[] builtRowFrom(T entity, List<Field> fieldsName) {
		Object[] ol = new Object[fieldsName.size()];
		int i = 0;
		for (Field field : fieldsName) {
			try {
				Object o = entity.getClass()
						.getMethod(
								"get" + Character.toUpperCase(field.getName().charAt(0)) + field.getName().substring(1))
						.invoke(entity);
				ol[i++] = o;
			} catch (IllegalAccessException | IllegalArgumentException | InvocationTargetException
					| NoSuchMethodException | SecurityException e) {
				e.printStackTrace();
			}
		}
		return ol;
	}
}
