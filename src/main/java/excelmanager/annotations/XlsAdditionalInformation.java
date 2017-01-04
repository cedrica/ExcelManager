/**
 * @author ca.leumaleu
 * 
 * Sometime there are overviews or footer information that have to be added in addition to the reports data.
 * This annotation allows you to set those additional information at a desired place of the sheet. 
 * Offers at the same time the possibility to merge the cells in which the text will be setted.
 */
package excelmanager.annotations;
 

import static java.lang.annotation.RetentionPolicy.RUNTIME;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.Target;

import excelmanager.enums.Location;

@Target({ ElementType.FIELD })
@Retention(RUNTIME)
public @interface XlsAdditionalInformation {
	String text() default "";
	Location location() default Location.TOP;
	int colspan() default 0;
}
