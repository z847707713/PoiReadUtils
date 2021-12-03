package cn.lovehao.poi.poiutils.excel.entity;

import cn.lovehao.poi.poiutils.excel.annotation.Excel;
import cn.lovehao.poi.poiutils.excel.enums.FieldType;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

@Data
public class TestExcel {

    @Excel(title = "商品图*",type = FieldType.IMG)
    private String img;

    @Excel(title = "商品名*")
    private String name;

    @Excel(title = "品类*")
    private String category;

    @Excel(title = "品牌*")
    private String brand;

    @Excel(title = "规格*")
    private String spec;

    @Excel(title = "备注")
    private String note;

    @Excel(title = "日期",type = FieldType.DATE)
    private Date time;

    @Excel(title = "数值",type = FieldType.NUMBER)
    private Double number;

    @Excel(title = "整型",type = FieldType.NUMBER)
    private Integer integerVal;

    @Excel(title = "长整型",type = FieldType.NUMBER)
    private Long longVal;

    @Excel(title = "decimal型",type = FieldType.NUMBER)
    private BigDecimal decimalVal;

    @Excel(title = "FORMULA",type = FieldType.NUMBER)
    private BigDecimal formula;

    private String emptyField;

}

