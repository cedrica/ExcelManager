package excelmanager.enums;

public enum Orientation {
	HORIZONTAL(0), OBLIC(45), VERTICAL(90);

	public final int value;

	Orientation(final int value) {
		this.value = value;
	}
}
