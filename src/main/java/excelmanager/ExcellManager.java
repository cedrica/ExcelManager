/**
 * this class is used to generate different kind of Excels report for PoJos
 * 
 *
 * @author ca.leumaleu
 */
package excelmanager;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import excelmanager.annotations.XLS;
import excelmanager.annotations.XlsColumn;
import excelmanager.enums.Declaratoin;
import excelmanager.enums.Location;
import excelmanager.enums.Orientation;
import excelmanager.exception.Assertion;
import excelmanager.helper.AdditionalInformation;
import excelmanager.style.HeaderStyleInfo;
import excelmanager.style.WoorkbookStyler;


@SuppressWarnings("all")
public class ExcellManager {
	private HashMap<String, HSSFCellStyle> hmFieldStyle = null;
//	private int rownum = 0;
//	private int cellnum = 0;
//	private Row row;
	
	public ExcellManager() {
	}

	/**
	 * generate a unique workbook for all POJOs inside the given list.
	 *
	 * @param entities
	 * @param entityClazz
	 */
	public <T> HSSFWorkbook generateSingleSheetReport(List<T> entities, Class entityClazz) {
		HSSFWorkbook workbook = new HSSFWorkbook();
		String sheetname = "";
		int rownum = 0;
		int cellnum = 0;
//		Row row;
		XLS xls = (XLS) entityClazz.getAnnotation(XLS.class);
		if (xls != null) {
			sheetname = (xls.sheetsname().trim().length() <= 0) ? entityClazz.getSimpleName() : xls.sheetsname();
		} else {
			return null;
		}
		HSSFSheet sheet = workbook.createSheet(sheetname);
		final List<Field> fields = Arrays.asList(entityClazz.getDeclaredFields());
		final List<Field> usedFields = new ArrayList<Field>();
		final List<String> nameOfUsedFields = new ArrayList<String>();
		hmFieldStyle = new HashMap<String, HSSFCellStyle>();
		WoorkbookStyler woorkbookStyler = new WoorkbookStyler();
		HeaderStyleInfo headerStyleInfo = new HeaderStyleInfo();
		fields.forEach(f -> {
			XlsColumn xlsAnnotation = f.getAnnotation(XlsColumn.class);
			if (xlsAnnotation != null) {
				String customname = xlsAnnotation.customname();
				short bgColor = xlsAnnotation.bgColor();
				short fgColor = xlsAnnotation.fgColor();
				Orientation orientation = xlsAnnotation.orientation();
				boolean italic = xlsAnnotation.isItalic();
				boolean bold = xlsAnnotation.isBold();
				String fontName = xlsAnnotation.fontName();
				int fontSize = xlsAnnotation.fontSize();

				headerStyleInfo.setOrientation(orientation);
				headerStyleInfo.setBgColor(bgColor);
				headerStyleInfo.setFgColor(fgColor);
				headerStyleInfo.setBold(HSSFFont.BOLDWEIGHT_BOLD);
				headerStyleInfo.setFontSize((short) fontSize);
				headerStyleInfo.setItalic(italic);
				headerStyleInfo.setFontName(fontName);

				woorkbookStyler.setHeaderStyleInfo(headerStyleInfo);
				woorkbookStyler.style(workbook);
				nameOfUsedFields.add((customname.length() <= 0) ? f.getName() : customname);
				usedFields.add(f);
				hmFieldStyle.put((customname.length() <= 0) ? f.getName() : customname, woorkbookStyler.headerStyle);
				System.out.println("custonname for field " + f.getName() + " => " + xlsAnnotation.customname());
			}
		});
		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
		data.put(0, nameOfUsedFields.toArray());
		int key = 1;
		for (T t : entities) {
			data.put(key++, builtRowFrom(t, usedFields));
		}

		Set<Integer> keyset = data.keySet();
		for (int k : keyset) {
			Row row = sheet.createRow(rownum);
			Object[] objArr = data.get(k);
			// Insert header line in the data
			if (rownum == 0) {
				for (Object obj : objArr) {
					Cell cell = row.createCell(cellnum++);
					cell.setCellValue(obj.toString());
					cell.setCellStyle(hmFieldStyle.get(obj.toString()));
				}
				rownum++;
				continue;
			}
			cellnum = 0;
			for (Object obj : objArr) {
				insertObjectAtRow(sheet, row, obj, cellnum);
			}
			rownum++;
		}
		return workbook;
//		if (fileToBeSaved != null) {
//			try {
//				File f = new File(fileToBeSaved.getPath());
//				if(f.isFile() && f.exists()){
//					f.delete();
//				}
//				FileOutputStream out = new FileOutputStream(f,false);
//				workbook.write(out);
//				out.close();
//				System.out.println("Excel written successfully..");
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
//			System.out.println("File saved: " + fileToBeSaved.getPath());
//		} else {
//			System.err.println("ERROR: File path is null.");
//		}
	}

	public void insertObjectAtRow(HSSFSheet sheet, Row row, Object obj, int cellnum) {
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
		cellnum++;
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
				// TODO Auto-generated catch block
				e.printStackTrace();
			}
		}

		return ol;
	}

	/**
	 * generate a single workbook for all POJOs inside the given list and add
	 * the additional information at the TOP or at the BOTTOM of the data
	 * content
	 *
	 * @param entities
	 * @param entityClazz
	 * @param additionalInformation
	 * @param location
	 */
	public <T> Workbook generateSingleSheetReportWithAdditionalInfo(List<T> entities, Class entityClazz,
			List<AdditionalInformation> additionalInformation, Location location) {
		HSSFWorkbook workbook = new HSSFWorkbook();
//		String sheetname = "";
//		XLS xls = (XLS) entityClazz.getAnnotation(XLS.class);
//		if (xls != null) {
//			sheetname = (xls.sheetsname().trim().length() <= 0) ? entityClazz.getSimpleName() : xls.sheetsname();
//		} else {
//			return workbook;
//		}
//		System.out.println("Sheetname = " + sheetname);
//		HSSFSheet sheet = workbook.createSheet(sheetname);
//		final List<Field> fields = Arrays.asList(entityClazz.getDeclaredFields());
//		final List<Field> usedFields = new ArrayList<Field>();
//		final List<String> nameOfUsedFields = new ArrayList<String>();
//		hmFieldStyle = new HashMap<String, HSSFCellStyle>();
//		WoorkbookStyler woorkbookStyler = new WoorkbookStyler();
//		HeaderStyleInfo headerStyleInfo = new HeaderStyleInfo();
//		fields.forEach(f -> {
//			XlsColumn xlsAnnotation = f.getAnnotation(XlsColumn.class);
//			if (xlsAnnotation != null) {
//				String customname = xlsAnnotation.customname();
//				short bgColor = xlsAnnotation.bgColor();
//				short fgColor = xlsAnnotation.fgColor();
//				Orientation orientation = xlsAnnotation.orientation();
//				boolean italic = xlsAnnotation.isItalic();
//				boolean bold = xlsAnnotation.isBold();
//				String fontName = xlsAnnotation.fontName();
//				int fontSize = xlsAnnotation.fontSize();
//
//				headerStyleInfo.setOrientation(orientation);
//				headerStyleInfo.setBgColor(bgColor);
//				headerStyleInfo.setFgColor(fgColor);
//				headerStyleInfo.setBold(HSSFFont.BOLDWEIGHT_BOLD);
//				headerStyleInfo.setFontSize((short) fontSize);
//				headerStyleInfo.setItalic(italic);
//				headerStyleInfo.setFontName(fontName);
//
//				woorkbookStyler.setHeaderStyleInfo(headerStyleInfo);
//				woorkbookStyler.style(workbook);
//				nameOfUsedFields.add((customname.length() <= 0) ? f.getName() : customname);
//				usedFields.add(f);
//				hmFieldStyle.put((customname.length() <= 0) ? f.getName() : customname, woorkbookStyler.headerStyle);
//				System.out.println("custonname for field " + f.getName() + " => " + xlsAnnotation.customname());
//			}
//		});
//		Map<Integer, Object[]> data = new HashMap<Integer, Object[]>();
//		rownum = 0;
//		cellnum = 0;
//		int key = 1;
//
//		boolean isHeaderAlreadyDone = false;
//		if (location == Location.TOP) {
//			int rowCount = additionalInformation.size();
//			additionalInformation.forEach(ids -> {
//				row = sheet.createRow(rownum++);
//				Cell input = row.createCell(cellnum++);
//				input.setCellValue(ids.getInput());
//				Cell value = row.createCell(cellnum++);
//				value.setCellValue(ids.getValue());
//				cellnum = 0;
//			});
//			row = sheet.createRow(rownum++);
//			lineSeparator(row, workbook, 2);
//			isHeaderAlreadyDone = true;
//		}
//		// Insert header line in the data
//		data.put(0, nameOfUsedFields.toArray());
//		for (T t : entities) {
//			data.put(key++, builtRowFrom(t, usedFields));
//		}
//		HSSFCellStyle rowStyle = workbook.createCellStyle();
//		rowStyle.setWrapText(true);
//		Set<Integer> keyset = data.keySet();
//		for (Integer k : keyset) {
//			row = sheet.createRow(rownum);
//			Object[] objArr = data.get(k);
//			cellnum = 0;
//			// Create header line in the excel doc
//			if (rownum == 0 || isHeaderAlreadyDone) {
//				for (Object obj : objArr) {
//					Cell cell = row.createCell(cellnum++);
//					cell.setCellValue(obj.toString());
//					cell.setCellStyle(hmFieldStyle.get(obj.toString()));
//				}
//				isHeaderAlreadyDone = false;
//				rownum++;
//				continue;
//			}
//			for (Object obj : objArr) {
//				Cell cell = row.createCell(cellnum);
//				if (obj instanceof Date) {
//					cell.setCellValue((Date) obj);
//				} else if (obj instanceof Boolean) {
//					cell.setCellValue((Boolean) obj);
//				} else if (obj instanceof String) {
//					cell.setCellValue((String) obj);
//				} else if (obj instanceof Double) {
//					cell.setCellValue((Double) obj);
//				} else if (obj instanceof Integer) {
//					cell.setCellValue((Integer) obj);
//				} else if (obj instanceof Long) {
//					cell.setCellValue((Long) obj);
//				} else if (obj instanceof Float) {
//					cell.setCellValue((Float) obj);
//				}
//				sheet.autoSizeColumn((short) cellnum);
//				cellnum++;
//			}
//			row.setRowStyle(rowStyle);
//			rownum++;
//		}
//
//		if (location == Location.BOTTOM) {
//			row = sheet.createRow(rownum++);
//			lineSeparator(row, workbook, 2);
//			int rowCount = additionalInformation.size();
//			cellnum = 0;
//			additionalInformation.forEach(ids -> {
//				row = sheet.createRow(rownum++);
//				Cell input = row.createCell(cellnum++);
//				input.setCellValue(ids.getInput());
//				Cell value = row.createCell(cellnum++);
//				value.setCellValue(ids.getValue());
//				cellnum = 0;
//			});
//		}

//		if (fileToBeSaved != null) {
//			try {
//				FileOutputStream out = new FileOutputStream(new File(fileToBeSaved.getPath()));
//				workbook.write(out);
//				out.close();
//				System.out.println("Excel written successfully..");
//			} catch (IOException e) {
//				e.printStackTrace();
//			}
//			System.out.println("File saved: " + fileToBeSaved.getPath());
//		} else {
//			System.out.println("File save cancelled.");
//		}
		return workbook;
	}

	private void lineSeparator(Row row, HSSFWorkbook workbook, int cellCount) {
		CellStyle cellStyle = workbook.createCellStyle();
		cellStyle.setFillForegroundColor(HSSFColor.GREY_50_PERCENT.index);
		cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		for (int i = 0; i < cellCount; i++) {
			row.createCell(i).setCellStyle(cellStyle);
			;
		}
	}

	/**
	 * Set a value in a sheet a the given cell position
	 *
	 * @param sheet
	 * @param rowIndex
	 * @param cellIndex
	 * @param value
	 */
	public void setValueAt(HSSFSheet sheet, long rowIndex, long cellIndex, Object value) {

		if (value instanceof Double) {
			sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellValue(Double.valueOf(value.toString()));
		} else if (value instanceof String) {
			sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellValue(value.toString());
		} else if (value instanceof Date) {
			sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellValue(value.toString());
		} else if (value instanceof Boolean) {
			sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellValue(Boolean.valueOf(value.toString()));
		} else if (value instanceof RichTextString) {
			sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellValue(value.toString());
		}

	}

	public void setCellFormular(HSSFCell cell, String formula) {
		cell.setCellFormula(formula);
	}

	public void setCellFormular(HSSFSheet sheet, long cellIndex, long rowIndex, String formula) {
		sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellFormula(formula);
	}

	
}
