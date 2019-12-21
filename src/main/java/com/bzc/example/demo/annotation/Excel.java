package com.bzc.example.demo.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Retention(RetentionPolicy.RUNTIME)
@Target({ElementType.TYPE})
public @interface Excel {

    String ALL = "_all_";

    /**
     * 标识标题读取起始行
     * @return
     */
    int skip() default 0;

    /**
     * 标识读取sheet页
     * @return
     */
    String[] sheets() default {"1"};
}