/**
 * This class create HSSFCellStyle object for header and data cells by mean of informations
 * setted in the style
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
	private StyleData styleData;

	public WoorkbookStyler() {
		styleData = new StyleData();
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
		if(styleData.isItalic()){
			
		}
		font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
		font.setItalic(styleData.isItalic());
		font.setFontName(styleData.getFontName());
		font.setFontHeightInPoints(styleData.getFontSize());
		font.setColor(styleData.getFgColor());
		style.setFont(font);
	}

	public void style(HSSFWorkbook wb) {
		style = wb.createCellStyle();
		createStyle(wb);
	}

	public void setStyle(StyleData styleData) {
		this.styleData = styleData;
	}

}
