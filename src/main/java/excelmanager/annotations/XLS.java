package excelmanager.annotations;


import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;
/**
 * Use this annotation on a class which has to compute into a excel report.
 * 
 * @author ca.leumaleu
 */
@Target({ElementType.TYPE})
@Retention(RUNTIME)
public @interface XLS {
	String sheetsname();
	XlsAdditionalInformation xlsAdditionalInformation() default @XlsAdditionalInformation();
}
