package io.velog.youmakemesmile.excel.config;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface HeaderStyle {
    Background background() default @Background;
    int fontSize() default 11;
    HorizontalAlignment horizontalAlignment() default HorizontalAlignment.CENTER;
    VerticalAlignment  verticalAlignment() default VerticalAlignment.CENTER;
}
