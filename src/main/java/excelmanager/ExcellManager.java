/**
 * this class is used to generate different kind of Excels report for PoJos
 * 
 *
 * @author ca.leumaleu
 */
package excelmanager;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
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
import org.apache.poi.ss.util.CellRangeAddress;
import org.dom4j.util.UserDataElement;

import excelmanager.annotations.XLS;
import excelmanager.annotations.XlsAdditionalInformation;
import excelmanager.annotations.XlsColumn;
import excelmanager.annotations.XlsStyler;
import excelmanager.enums.Location;
import excelmanager.enums.Orientation;
import excelmanager.style.StyleData;
import excelmanager.style.WoorkbookStyler;


@SuppressWarnings("all")
public class ExcellManager {
	private HashMap<String, HSSFCellStyle> hmFieldStyle = null;
	private static int FRIST_ROW_INDEX_FOR_DATA = 0;
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
		HashMap<Location, String> footerInfo = new HashMap<>();
		Row row;
		XLS xls = (XLS) entityClazz.getAnnotation(XLS.class);
		XlsAdditionalInformation additionalInformation = null;
		if (xls != null) {
			sheetname = (xls.sheetsname().trim().length() <= 0) ? entityClazz.getSimpleName() : xls.sheetsname();
			additionalInformation = xls.xlsAdditionalInformation();
		} else {
			return null;
		}
		HSSFSheet sheet = workbook.createSheet(sheetname);
		if(additionalInformation != null){
			String text = additionalInformation.text();
			Location location = additionalInformation.location();
			int colspan = additionalInformation.colspan();
			if(location == Location.BOTTOM){
				footerInfo.put(Location.BOTTOM,text+" colspan= "+colspan);
			}else {
				row = sheet.createRow(rownum++);
				FRIST_ROW_INDEX_FOR_DATA = rownum;
				Cell cell = row.createCell((short)0);
				cell.setCellValue(text);
				sheet.addMergedRegion(new CellRangeAddress(
		                        0, // mention first row here
		                        0, //mention last row here, it is 1 as we are doing a column wise merging
		                        0, //mention first column of merging
		                        colspan  //mention last column to include in merge
		                        ));	
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
				StyleData styleData = new StyleData();
				String customname = xlsAnnotation.customname();
				XlsStyler xlsStyler = xlsAnnotation.styler();
				styleData = extractStyleInfo(xlsStyler);
				woorkbookStyler.setStyle(styleData);
				woorkbookStyler.style(workbook);
				nameOfUsedFields.add((customname.length() <= 0) ? f.getName() : customname);
				usedFields.add(f);
				hmFieldStyle.put((customname.length() <= 0) ? f.getName() : customname, woorkbookStyler.style);
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
			row = sheet.createRow(rownum);
			Object[] objArr = data.get(k);
			// Insert header line in the data
			if (rownum == FRIST_ROW_INDEX_FOR_DATA) {
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
				cellnum++;
			}
			rownum++;
		}
		if(footerInfo != null &&  footerInfo.size() > 0){
			String textColspan = footerInfo.get(Location.BOTTOM);
			if (textColspan.trim().length() > 0){
				String[] splittedStr = textColspan.split("colspan=");
				String text = splittedStr[0];
				int colspan = Integer.valueOf(splittedStr[1].trim());
				row = sheet.createRow(rownum++);
				Cell cell = row.createCell((short)0);
				cell.setCellValue(text);
				sheet.addMergedRegion(new CellRangeAddress(
								entities.size()+1, // mention first row here
								entities.size()+1, //mention last row here, it is 1 as we are doing a column wise merging
		                        0, //mention first column of merging
		                        colspan  //mention last column to include in merge
		                        ));	
			}

		}
		return workbook;
	}

	public StyleData extractStyleInfo( XlsStyler xlsStyler) {
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
		styleData.setBold(HSSFFont.BOLDWEIGHT_BOLD);
		styleData.setFontSize((short) fontSize);
		styleData.setItalic(italic);
		styleData.setFontName(fontName);
		return styleData;
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

	public void setCellFormular(HSSFCell cell, String formula) {
		cell.setCellFormula(formula);
	}

	public void setCellFormular(HSSFSheet sheet, long cellIndex, long rowIndex, String formula) {
		sheet.getRow((int) rowIndex).getCell((int) cellIndex).setCellFormula(formula);
	}

	
}
