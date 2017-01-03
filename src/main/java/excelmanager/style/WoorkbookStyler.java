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

	public static HSSFCellStyle headerStyle;
	public static HSSFCellStyle cellStyle;
	private HeaderStyleInfo headerStyleInfo;
	private CellStyleInfo cellStyleInfo;

	public WoorkbookStyler() {
		headerStyleInfo = new HeaderStyleInfo();
		cellStyleInfo = new CellStyleInfo();
	}

	public void createHeaderStyle(HSSFWorkbook wb) {
		HSSFFont font = wb.createFont();

		if (headerStyleInfo.getFormat() != null)
			headerStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(headerStyleInfo.getFormat().toString()));

		if (headerStyleInfo.getAlignment() != null) {
			headerStyle.setAlignment(headerStyleInfo.getAlignment().shortValue());
		}

		if (headerStyleInfo.getOrientation() != null) {
			headerStyle.setRotation((short) headerStyleInfo.getOrientation().value);
		}
		headerStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		headerStyle.setFillForegroundColor(headerStyleInfo.getBgColor());
		font.setBoldweight(headerStyleInfo.getBold());
		font.setItalic(headerStyleInfo.isItalic());
		font.setFontName(headerStyleInfo.getFontName());
		font.setFontHeightInPoints(headerStyleInfo.getFontSize());
		font.setColor(headerStyleInfo.getFgColor());
		headerStyle.setFont(font);
	}


	public void createDataCellStyle(HSSFWorkbook wb) {
		HSSFFont font = wb.createFont();
		font = wb.createFont();
		if (cellStyleInfo.getFormat() != null)
			cellStyle.setDataFormat(HSSFDataFormat.getBuiltinFormat(cellStyleInfo.getFormat().toString()));
		if (cellStyleInfo.getAlignment() != null) {
			cellStyle.setAlignment(cellStyleInfo.getAlignment().shortValue());
		}
		if (cellStyleInfo.getOrientation() != null) {
			cellStyle.setRotation((short) cellStyleInfo.getOrientation().value);
		}
		if (cellStyleInfo.getBgColor() != null) {
			cellStyle.setFillBackgroundColor(cellStyleInfo.getBgColor());
		}
		if (cellStyleInfo.getFontColor() != null) {
			cellStyle.setFillForegroundColor(cellStyleInfo.getFontColor());
		}
		if (cellStyleInfo.getFontSize() != null) {
			font.setFontHeight(cellStyleInfo.getFontSize());
		}
		if (cellStyleInfo.getFontStyle() != null) {
			font.setBoldweight(cellStyleInfo.getFontStyle());
		}
		cellStyle.setFont(font);

	}
	public void style(HSSFWorkbook wb) {
		headerStyle = wb.createCellStyle();
		cellStyle = wb.createCellStyle();
		createHeaderStyle(wb);
		createDataCellStyle(wb);
	}

	public HeaderStyleInfo getHeaderStyleInfo() {
		return headerStyleInfo;
	}
	
	public void setHeaderStyleInfo(HeaderStyleInfo headerStyleInfo) {
		this.headerStyleInfo = headerStyleInfo;
	}

	public CellStyleInfo getCellStyleInfo() {
		return cellStyleInfo;
	}

	public void setCellStyleInfo(CellStyleInfo cellStyleInfo) {
		this.cellStyleInfo = cellStyleInfo;
	}

}
