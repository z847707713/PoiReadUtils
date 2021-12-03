package cn.lovehao.poi.poiutils.excel.annotation;

import cn.lovehao.poi.poiutils.excel.enums.FieldType;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 *  excel 的注解类
 */

@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface Excel {

    /**
     *  表格的列名
     * @return
     */
    String title();

    FieldType type() default FieldType.STRING;

}
