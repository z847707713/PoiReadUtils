package cn.lovehao.poi.poiutils.excel.utils;

import cn.lovehao.poi.poiutils.excel.annotation.Excel;
import cn.lovehao.poi.poiutils.excel.enums.FieldType;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPicture;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.springframework.beans.factory.InitializingBean;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.util.*;

@Component
public class PoiReadUtils implements InitializingBean {

    @Value("${img.path}")
    private String iconPath;

    private static String IMG_PATH;


    public static <T> List<T> read(String path, Class<T> clazz){
        File file = new File(path);
        return read(file,clazz);
    }


    public static <T> List<T> read(File file, Class<T> clazz){
        if(!file.exists()){
            throw new RuntimeException("文件不存在");
        }

        Workbook workbook = null;
        try {
            workbook = WorkbookFactory.create(file);
            Sheet sheet = workbook.getSheetAt(0);
            FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();


            // 图片相关处理
            Drawing<?> drawingPatriarch = sheet.getDrawingPatriarch();
            Map<String, PictureData> pictureDataMap = new HashMap<>();
            for (Object item : drawingPatriarch) {
                if (item instanceof HSSFPicture) {
                    HSSFPicture hssfPicture = (HSSFPicture) item;
                    HSSFClientAnchor hssfClientAnchor = (HSSFClientAnchor) hssfPicture.getAnchor();
                    PictureData pictureData = hssfPicture.getPictureData();
                    String key = hssfClientAnchor.getRow1() + "-" + hssfClientAnchor.getCol1();
                    pictureDataMap.put(key, pictureData);
                }
            }

            Field[] fields = clazz.getDeclaredFields();

            // 列名与字段映射
            Map<String, Field> fieldMap = new HashMap<>();
            // 列名与注解映射
            Map<String, Excel> annotationMap = new HashMap<>();

            for (Field field : fields) {
                Excel excel = field.getAnnotation(Excel.class);
                if (excel != null) {
                    fieldMap.put(excel.title(), field);
                    annotationMap.put(excel.title(), excel);
                }
            }

            Map<Integer, String> cellIndexMap = new HashMap<>();

            Row titleRow = sheet.getRow(0);

            titleRow.forEach( cell -> cellIndexMap.put(cell.getColumnIndex(), cell.getRichStringCellValue().getString()));

            List<T> result = new ArrayList<>();

            for (Row row : sheet) {
                if (row.getRowNum() == 0) {
                    continue;
                }
                try {
                    T obj = clazz.newInstance();

                    for (Cell cell : row) {
                        String title = cellIndexMap.get(cell.getColumnIndex());

                        // 如果对象中不包含这个列，则跳过
                        if (!annotationMap.containsKey(title)) {
                            continue;
                        }

                        Excel annotation = annotationMap.get(title);
                        Field field = fieldMap.get(title);
                        field.setAccessible(true);

                        switch (cell.getCellType()) {
                            case STRING:
                                setStringVal(obj, cell, annotation, field);
                                break;
                            case NUMERIC:
                                setNumberVal(obj, cell, annotation, field);
                                break;
                            case BOOLEAN:
                                setBooleanVal(obj, cell, annotation, field);
                                break;
                            case FORMULA:
                                CellValue cellValue = evaluator.evaluate(cell);
                                switch (cellValue.getCellType()){
                                    case STRING:
                                        setStringVal(obj, annotation, field, cellValue);
                                        break;
                                    case BOOLEAN:
                                        setBooleanVal(obj, annotation, field, cellValue);
                                        break;
                                    case NUMERIC:
                                        setNumberVal(workbook, obj, annotation, field, cellValue);
                                        break;
                                    case BLANK:
                                        setStringVal(obj, annotation, field, cellValue);
                                        break;
                                    default:
                                }
                                break;
                            case BLANK:
                                setImgVal(pictureDataMap, row, cell, obj , annotation, field);
                                break;
                            default:
                                System.out.println("行:" + cell.getRowIndex() + ",列:" + cell.getColumnIndex() + ".默认");
                        }
                    }

                    result.add(obj);

                } catch (InstantiationException | IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
            return result;

        } catch (IOException e) {
            e.printStackTrace();
        }
        return null;
    }


    private static <T> void setImgVal(Map<String, PictureData> pictureDataMap, Row row, Cell cell, T obj , Excel annotation, Field field) {
        if(annotation.type() == FieldType.IMG && String.class.isAssignableFrom(field.getType())){

            String key = row.getRowNum() + "-" + cell.getColumnIndex();
            if(pictureDataMap.containsKey(key)){
                PictureData pictureData = pictureDataMap.get(key);
                byte[] bytes = pictureData.getData();

                String type = ".png";
                switch (pictureData.getPictureType()){
                    case 5:
                        type = ".png";
                        break;
                    case 6:
                        type = ".jpeg";
                        break;
                }

                String relativeFilePath = "\\" + UUID.randomUUID() + type;

                File imgFile = new File(IMG_PATH + relativeFilePath);
                if(!imgFile.exists()){
                    try {
                        imgFile.createNewFile();
                        try( FileOutputStream fileOutputStream = new FileOutputStream(imgFile)){
                            fileOutputStream.write(bytes);
                        }
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                try {
                    field.set(obj,relativeFilePath);
                    System.out.println(field.get(obj));
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }

            }
        }
    }

    private static <T> void setNumberVal(Workbook workbook, T obj, Excel annotation, Field field, CellValue cellValue) {
        if (annotation.type() == FieldType.DATE && field.getType().equals(Date.class)) {
            try {
                if(HSSFWorkbook.class.isAssignableFrom(workbook.getClass())){
                    HSSFWorkbook hssfWorkbook = (HSSFWorkbook)workbook;
                    double val = cellValue.getNumberValue();
                    Date date = hssfWorkbook.getInternalWorkbook().isUsing1904DateWindowing() ? DateUtil.getJavaDate(val, true) : DateUtil.getJavaDate(val, false);
                    field.set(obj,  date);
                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
        if (annotation.type() == FieldType.NUMBER && Number.class.isAssignableFrom(field.getType())) {
            try {
                if(field.getType().isAssignableFrom(Double.class)){
                    field.set(obj,  cellValue.getNumberValue());
                } else if(field.getType().isAssignableFrom(BigDecimal.class)){
                    field.set(obj,  BigDecimal.valueOf(cellValue.getNumberValue()));
                } else if(field.getType().isAssignableFrom(Integer.class)){
                    field.set(obj, Double.valueOf(cellValue.getNumberValue()).intValue());
                } else if(field.getType().isAssignableFrom(Long.class)){
                    field.set(obj, Double.valueOf(cellValue.getNumberValue()).longValue());
                }
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    private static <T> void setBooleanVal(T obj, Excel annotation, Field field, CellValue cellValue) {
        if (annotation.type() == FieldType.BOOLEAN && Boolean.class.isAssignableFrom(field.getType())) {
            try {
                field.set(obj,cellValue.getBooleanValue());
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    private static <T> void setStringVal(T obj, Excel annotation, Field field, CellValue cellValue) {
        if (annotation.type() == FieldType.STRING && String.class.isAssignableFrom(field.getType())) {
            try {
                field.set(obj, cellValue.getStringValue());
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    private static <T> void setBooleanVal(T obj, Cell cell, Excel annotation, Field field) {
        if (annotation.type() == FieldType.NUMBER && Boolean.class.isAssignableFrom(field.getType())) {
            try {
                field.set(obj,cell.getBooleanCellValue());
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }

    private static <T> void setNumberVal(T obj, Cell cell, Excel annotation, Field field) {
        if (DateUtil.isCellDateFormatted(cell)) {
            if (annotation.type() == FieldType.DATE && field.getType().equals(Date.class)) {
                try {
                    field.set(obj,  cell.getDateCellValue());
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        } else {
            if (annotation.type() == FieldType.NUMBER && Number.class.isAssignableFrom(field.getType())) {
                try {
                    if(field.getType().isAssignableFrom(Double.class)){
                        field.set(obj,  cell.getNumericCellValue());
                    } else if(field.getType().isAssignableFrom(BigDecimal.class)){
                        field.set(obj,  BigDecimal.valueOf(cell.getNumericCellValue()));
                    } else if(field.getType().isAssignableFrom(Integer.class)){
                        field.set(obj, Double.valueOf(cell.getNumericCellValue()).intValue());
                    } else if(field.getType().isAssignableFrom(Long.class)){
                        field.set(obj, Double.valueOf(cell.getNumericCellValue()).longValue());
                    }
                } catch (IllegalAccessException e) {
                    e.printStackTrace();
                }
            }
        }
    }

    private static <T> void setStringVal(T obj, Cell cell, Excel annotation, Field field) {
        String val = cell.getRichStringCellValue().getString();
        if (annotation.type() == FieldType.STRING && field.getType().equals(String.class)) {
            try {
                field.set(obj, val);
            } catch (IllegalAccessException e) {
                e.printStackTrace();
            }
        }
    }


    @Override
    public void afterPropertiesSet() throws Exception {
        IMG_PATH = iconPath;
    }
}
