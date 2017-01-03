package excelmanager.annotations;


import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;
/**
 * Annotate a class with this annotation to declare it as a report
 * 
 * @author ca.leumaleu
 */
@Target({ElementType.TYPE})
@Retention(RUNTIME)
public @interface XLS {
	String sheetsname();
	XlsAdditionalInformation xlsAdditionalInformation() default @XlsAdditionalInformation();
}
