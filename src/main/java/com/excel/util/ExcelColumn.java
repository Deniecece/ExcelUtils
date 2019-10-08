package com.excel.util;

import java.lang.annotation.*;

@Target({ElementType.FIELD})
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {

    /**
     * 对应excel的列名
     * @return
     */
    String value() default "";

    /**
     * 列数排行
     * @return
     */
    int col() default 0;
}
