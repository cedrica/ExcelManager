/**
 * @author ca.leumaleu
 * 
 * This annotation is used on fields of a PoJos that must be take in consideration during the generation of the excel-report.
 * If a field is not annotated then it will not appears in the report.
 */
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

}
