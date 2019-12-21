package com.bzc.example.demo.annotation;

import com.bzc.example.demo.utils.Type;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.FIELD})
public @interface ExcelFied {

    //列索引
    int index() default 0;

    //关联列数据查询语句(类枚举类型)
    String sqlStr() default "";

    boolean isTrim() default false;

    /**
     *  是否允许为空,默认允许false
     * @return
     */
    boolean notNull() default false;

    /**
     *  长度
     * @return
     */
    int length() default 0;

    /**
     *  类型
     * @return
     */
    Type type() default Type.STRING;

    /**
     *  日期格式
     * @return
     */
    String dateFormatter() default "";

    /**
     *  最小值
     * @return
     */
    String minValue() default "";

    /**
     *  最大值
     * @return
     */
    String maxValue() default "";
}