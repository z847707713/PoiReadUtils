package cn.lovehao.poi.poiutils.excel.enums;

public enum FieldType {

    STRING(0,"字符串"),
    NUMBER(1,"数值"),
    DATE(2,"日期"),
    IMG(3,"图片"),
    BOOLEAN(4,"布尔值"),
    ;

    FieldType(int typeCd, String typeName) {
        this.typeCd = typeCd;
        this.typeName = typeName;
    }

    private int typeCd;

    private String typeName;


}
