package com.moer.poi.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target({ElementType.FIELD})//可以定义在方法上
@Retention(RetentionPolicy.RUNTIME)//运行有效,存在class字节码文件中
public @interface MyExcelAnno {
    int type() default 0; //0标准格式 1带下拉框 2带格式化
    int soft() default 0;	//下拉框排序
    String value() default "";//列名的中文
    String format() default "";//要输出的格式用 "{}万"
}
