package excelmanager.style;

import excelmanager.enums.Format;
import excelmanager.enums.Orientation;

public class StyleData {
	private Short alignment;
	private Short bgColor;
	private Short fontColor;
	private Short fontSize;
	private Short fontStyle;
	private Format format;
	private Orientation orientation;
	private short fgColor;
	private boolean italic;
	private short bold;
	private String fontName;

	public StyleData() {
	}

	public Orientation getOrientation() {
		return orientation;
	}


	public void setOrientation(Orientation orientation) {
		this.orientation = orientation;
	}


	public Short getAlignment() {
		return alignment;
	}

	public void setAlignment(Short alignment) {
		this.alignment = alignment;
	}


	public Short getBgColor() {
		return bgColor;
	}


	public void setBgColor(Short bgColor) {
		this.bgColor = bgColor;
	}


	public Short getFontColor() {
		return fontColor;
	}


	public void setFontColor(Short fontColor) {
		this.fontColor = fontColor;
	}

	public Short getFontSize() {
		return fontSize;
	}


	public void setFontSize(Short fontSize) {
		this.fontSize = fontSize;
	}


	public Short getFontStyle() {
		return fontStyle;
	}


	public void setFontStyle(Short fontStyle) {
		this.fontStyle = fontStyle;
	}


	public Format getFormat() {
		return format;
	}


	public void setFormat(Format format) {
		this.format = format;
	}


	public void setFgColor(short fgColor) {
		this.fgColor = fgColor;
	}
	public short getFgColor() {
		return this.fgColor;
	}


	public void setItalic(boolean itatlic) {
		this.italic = itatlic;
	}

	public boolean isItalic() {
		return this.italic;
	}

	public void setBold(short bold) {
		this.bold = bold;
	}

	public short getBold() {
		return this.bold;
	}

	public String getFontName() {
		return this.fontName;
	}
	public void setFontName(String fontName) {
		this.fontName= fontName;
	}
}
