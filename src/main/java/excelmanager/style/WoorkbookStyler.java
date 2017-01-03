/**
 * This class create HSSFCellStyle object for header and data cells by mean of informations
 * setted in the headerStyleInfo and cellStyleInfo
 *
 * @author ca.leumaleu
 */
package excelmanager.style;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDataFormat;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class WoorkbookStyler {

	public static HSSFCellStyle style;
//	public static HSSFCellStyle cellStyle;
	private StyleData styleData;
	private CellStyleInfo cellStyleInfo;

	public WoorkbookStyler() {
		styleData = new StyleData();
//		cellStyleInfo = new CellStyleInfo();
	}

	public void createStyle(HSSFWorkbook wb) {
		HSSFFont font = wb.createFont();

		if (styleData.getFormat() != null)
			style.setDataFormat(HSSFDataFormat.getBuiltinFormat(styleData.getFormat().toString()));

		if (styleData.getAlignment() != null) {
			style.setAlignment(styleData.getAlignment().shortValue());
		}

		if (styleData.getOrientation() != null) {
			style.setRotation((short) styleData.getOrientation().value);
		}
		style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		style.setFillForegroundColor(styleData.getBgColor());
		font.setBoldweight(styleData.getBold());
		font.setItalic(styleData.isItalic());
		font.setFontName(styleData.getFontName());
		font.setFontHeightInPoints(styleData.getFontSize());
		font.setColor(styleData.getFgColor());
		style.setFont(font);
	}


//	public void createDataCellStyle(HSSFWorkbook wb) {
//		HSSFFont font = wb.createFont();
//		font = wb.createFont();
//		if (cellStyleInfo.getFormat() != null)
//			cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(cellStyleInfo.getFormat().toString()));
//		if (cellStyleInfo.getAlignment() != null) {
//			cellStyle.setAlignment(cellStyleInfo.getAlignment().shortValue());
//		}
//		if (cellStyleInfo.getOrientation() != null) {
//			cellStyle.setRotation((short) cellStyleInfo.getOrientation().value);
//		}
//		if (cellStyleInfo.getBgColor() != null) {
//			cellStyle.setFillBackgroundColor(cellStyleInfo.getBgColor());
//		}
//		if (cellStyleInfo.getFontColor() != null) {
//			cellStyle.setFillForegroundColor(cellStyleInfo.getFontColor());
//		}
//		if (cellStyleInfo.getFontSize() != null) {
//			font.setFontHeight(cellStyleInfo.getFontSize());
//		}
//		if (cellStyleInfo.getFontStyle() != null) {
//			font.setBoldweight(cellStyleInfo.getFontStyle());
//		}
//		cellStyle.setFont(font);
//
//	}
	public void style(HSSFWorkbook wb) {
		style = wb.createCellStyle();
//		cellStyle = wb.createCellStyle();
		createStyle(wb);
//		createDataCellStyle(wb);
	}

	public void setStyle(StyleData styleData) {
		this.styleData = styleData;
	}



}
