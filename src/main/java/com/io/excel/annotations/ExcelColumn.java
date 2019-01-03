package com.io.excel.annotations;

import java.lang.annotation.Documented;
import java.lang.annotation.ElementType;
import java.lang.annotation.Inherited;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;
import java.util.Comparator;

/**
 *
 * @author pcoelho
 */
@Documented
@Inherited
@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelColumn {
    String name() default "";
    String[] columnDefinitions() default "";
    String index() default "";
    String subclassField() default "";
    // TODO: String validationMethod() default "";
    // TODO: boolean interruptiveValidation() default false;
    
    class ColumnComparator implements Comparator<String>
    {
        @Override
        public int compare(String o1, String o2) {             
            if (o1.length()!=o2.length()) {
                return o1.length()-o2.length(); //overflow impossible since lengths are non-negative
            }
            return o1.compareTo(o2);
        }
    }
}
