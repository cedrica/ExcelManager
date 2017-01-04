/**
 * @author ca.leumaleu
 * use for cells styling
 */
package excelmanager.annotations;

import static java.lang.annotation.ElementType.FIELD;

import java.lang.annotation.Retention;
import static java.lang.annotation.RetentionPolicy.RUNTIME;
import java.lang.annotation.Target;

import org.apache.poi.hssf.util.HSSFColor;

import excelmanager.enums.Orientation;


@Retention(RUNTIME)
@Target({FIELD})
public @interface XlsStyler {
	public Orientation orientation() default Orientation.HORIZONTAL;

	public short bgColor() default HSSFColor.WHITE.index;

	public short fgColor() default HSSFColor.BLACK.index;

	public int fontSize() default 10;

	public String fontName() default "Arial";

	public boolean isBold() default false;

	public boolean isItalic() default false;
}
