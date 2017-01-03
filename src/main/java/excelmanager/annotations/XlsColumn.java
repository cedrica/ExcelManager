package excelmanager.annotations;

import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

@Target({ ElementType.FIELD })
@Retention(RUNTIME)
public @interface XlsColumn {

	public String customname() default "";
	
	public XlsStyler styler() default @XlsStyler;


//	public Orientation orientation() default Orientation.HORIZONTAL;
//
//	public short bgColor() default HSSFColor.WHITE.index;
//
//	public short fgColor() default HSSFColor.BLACK.index;
//
//	public int fontSize() default 10;
//
//	public String fontName() default "Arial";
//
//	public boolean isBold() default false;
//
//	public boolean isItalic() default false;

}
