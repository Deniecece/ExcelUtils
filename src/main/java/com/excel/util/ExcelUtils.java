package com.excel.util;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.MediaType;
import org.apache.commons.lang3.BooleanUtils;
import org.apache.commons.lang3.CharUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.commons.lang3.math.NumberUtils;
import org.apache.commons.collections.CollectionUtils;
import org.springframework.lang.Nullable;
import org.springframework.web.context.request.RequestContextHolder;
import org.springframework.web.context.request.ServletRequestAttributes;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.Null;
import java.io.*;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.math.BigDecimal;
import java.net.URLEncoder;
import java.util.*;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * excel读写工具类
 *
 * @author shantong
 * @description excel utils
 * @date 2018/9/12 15:27
 **/
public class ExcelUtils {

    private final static Logger log = LoggerFactory.getLogger(ExcelUtils.class);

    private final static String EXCEL2003 = "xls";
    private final static String EXCEL2007 = "xlsx";

    /**
     * 读取excel为对应的实体类
     * @param cls 实体类
     * @param file excel文件
     * @return
     */
    public static <T> List<T> readExcelObject(Class<T> cls, MultipartFile file){

        //1.文件格式验证
        String fileName = file.getOriginalFilename();
        if (!fileName.matches("^.+\\.(?i)(xls)$") && !fileName.matches("^.+\\.(?i)(xlsx)$")) {
            log.error("上传文件格式不正确");
        }
        List<T> dataList = new ArrayList<>();
        Workbook workbook = null;
        try {

            //2.创建对应版本的wb
            workbook = getWorkbook(file.getInputStream(), fileName);
            if (workbook != null) {

                //映射  注解的value(excel的列名)-对应->实体类的属性
                Map<String, List<Field>> classMap = new HashMap<>();
                List<Field> fields = Stream.of(cls.getDeclaredFields()).collect(Collectors.toList());
                fields.forEach(
                        field -> {
                            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                            if (annotation != null) {
                                String value = annotation.value();
                                if (StringUtils.isBlank(value)) {
                                    return;
                                }
                                if (!classMap.containsKey(value)) {
                                    classMap.put(value, new ArrayList<>());
                                }
                                field.setAccessible(true);
                                classMap.get(value).add(field);
                            }
                        }
                );

                //索引Map-->excel的列名序号 对应 字段
                Map<Integer, List<Field>> reflectionMap = new HashMap<>(16);

                Sheet sheet = workbook.getSheetAt(0);
                boolean firstRow = true;
                for (int i = sheet.getFirstRowNum(); i <= sheet.getLastRowNum(); i++) {
                    Row row = sheet.getRow(i);
                    //处理索引Map
                    if (firstRow) {
                        for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                            Cell cell = row.getCell(j);
                            String cellValue = getCellValue(cell);
                            if (classMap.containsKey(cellValue)) {
                                reflectionMap.put(j, classMap.get(cellValue));
                            }
                        }
                        firstRow = false;
                    } else {
                        //空白行
                        if (row == null) {
                            continue;
                        }
                        try {
                            T bean = cls.newInstance();
                            //空白行
                            boolean allBlank = true;
                            for (int j = row.getFirstCellNum(); j <= row.getLastCellNum(); j++) {
                                if (reflectionMap.containsKey(j)) {
                                    Cell cell = row.getCell(j);
                                    String cellValue = getCellValue(cell);
                                    if (StringUtils.isNotBlank(cellValue)) {
                                        allBlank = false;
                                    }
                                    List<Field> fieldList = reflectionMap.get(j);
                                    fieldList.forEach(
                                            field -> {
                                                try {
                                                    //字段赋值
                                                    fillField(bean, cellValue, field);
                                                } catch (Exception e) {
                                                    log.error(String.format("字段:%s赋值:%s 异常!", field.getName(), cellValue), e);
                                                }
                                            }
                                    );
                                }
                            }
                            if (!allBlank) {
                                dataList.add(bean);
                            } else {
                                log.warn(String.format("第:%s 行为空!", i));
                            }
                        } catch (Exception e) {
                            log.error(String.format("行:%s 转化异常!", i), e);
                        }
                    }
                }
            }
        } catch (Exception e) {
            log.error(String.format("excel 转化异常!"), e);
        } finally {
            if (workbook != null) {
                try {
                    workbook.close();
                } catch (Exception e) {
                    log.error(String.format("excel 转化异常!"), e);
                }
            }
        }
        return dataList;
    }

    /**
     * 读取excel指定sheet的内容为Map
     * @param in
     * @param fileName
     * @param sheetIndex
     * @return
     */
    public static List<Map<String, String>> readExcelMap(InputStream in, String fileName, int sheetIndex){
        List<Map<String, String>> list = new ArrayList<>();
        Workbook work = getWorkbook(in, fileName);
        if (null == work || work.getSheetAt(sheetIndex) == null) {
            log.error("上传Excel中Sheet:{}为空", sheetIndex);
        }
        list.addAll(readSheetList(work.getSheetAt(sheetIndex)));
        return list;
    }

    /**
     * 读取excel单个sheet的集合
     * @param sheet
     * @return
     */
    public static List<Map<String, String>> readSheetList(Sheet sheet) {
        List<Map<String, String>> rows = new ArrayList<>();
        Row fristRow = sheet.getRow(0);
        Row row;
        List<String> values;
        int lastRowNum = sheet.getLastRowNum();
        for (int i = 1; i <= lastRowNum; i++) {
            row = sheet.getRow(i);
            if (row == null) {
                break;
            }
            values = new ArrayList();
            for (int j = 0; j < row.getLastCellNum(); j++) {
                values.add(getCellValue(row.getCell(j)));
            }
            if (CollectionUtils.isEmpty(values)) {
                break;
            }
            Map<String, String> rowMap = rowToMap(fristRow, values);
            if (rowMap != null) {
                rows.add(rowMap);
            }
        }
        return rows;
    }

    /**
     * 读取单元值
     * @param cell
     * @return
     */
    public static String getCellValue(Cell cell) {
        String value = "";
        if (cell == null) {
            return value;
        }
        switch (cell.getCellTypeEnum()) {
            case NUMERIC :
                value = DateUtil.isCellDateFormatted(cell) ? HSSFDateUtil.getJavaDate(cell.getNumericCellValue()).toString() :
                        new BigDecimal(cell.getNumericCellValue()).toString();
                break;
            case STRING :
                value = StringUtils.trimToEmpty(cell.getStringCellValue());
                break;
            case FORMULA :
                value = StringUtils.trimToEmpty(cell.getCellFormula());
                break;
            case BOOLEAN :
                value = String.valueOf(cell.getBooleanCellValue());
                break;
            case BLANK:
                value = "";
                break;
            case ERROR :
                value = "ERROR";
                break;
            default:
                value = cell.toString().trim();
                break;
        }
        return value;
    }

    /**
     * 实体类集合写excel
     * 通过在实体类添加注解实现与excel列名对应
     * @param dataList
     * @param cls
     * @param <T>
     */
    public static <T> void writeExcelByAnnotation(List<T> dataList, Class<T> cls, HttpServletResponse response){
        Field[] fields = cls.getDeclaredFields();
        List<Field> fieldList = Arrays.stream(fields)
                .filter(field -> {
                    ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                    if (annotation != null && annotation.col() > 0) {
                        field.setAccessible(true);
                        return true;
                    }
                    return false;
                }).sorted(Comparator.comparing(field -> {
                    int col = 0;
                    ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                    if (annotation != null) {
                        col = annotation.col();
                    }
                    return col;
                })).collect(Collectors.toList());

        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        AtomicInteger ai = new AtomicInteger();
        {
            Row row = sheet.createRow(ai.getAndIncrement());
            AtomicInteger aj = new AtomicInteger();
            //写入头部
            fieldList.forEach(field -> {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                String columnName = "";
                if (annotation != null) {
                    columnName = annotation.value();
                }
                Cell cell = row.createCell(aj.getAndIncrement());

                CellStyle cellStyle = wb.createCellStyle();
                cellStyle.setFillForegroundColor(IndexedColors.WHITE.getIndex());
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setAlignment(HorizontalAlignment.CENTER);

//                Font font = wb.createFont();
//                font.ssetBoldweight(Font.BOLDWEIGHT_NORMAL);
//                cellStyle.setFont(font);
                cell.setCellStyle(cellStyle);
                cell.setCellValue(columnName);
            });
        }
        if (CollectionUtils.isNotEmpty(dataList)) {
            dataList.forEach(t -> {
                Row row1 = sheet.createRow(ai.getAndIncrement());
                AtomicInteger aj = new AtomicInteger();
                fieldList.forEach(field -> {
                    Class<?> type = field.getType();
                    Object value = "";
                    try {
                        value = field.get(t);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    Cell cell = row1.createCell(aj.getAndIncrement());
                    if (value != null) {
                        if (type == Date.class) {
                            cell.setCellValue(value.toString());
                        } else {
                            cell.setCellValue(value.toString());
                        }
                        cell.setCellValue(value.toString());
                    }
                });
            });
        }
        //冻结窗格
        wb.getSheet("Sheet1").createFreezePane(0, 1, 0, 1);
        //浏览器下载excel
        downloadExcel("abbot.xlsx",wb, response);
        //生成excel文件
//        buildExcelFile(".\\default.xlsx",wb);
    }

    /**
     * 实体类集合写excel
     * 没有使用注解需要传递转化的字段fields
     * @param objects 需要转化为Excel的对象
     * @param fields  需要转化的字段
     * @param titles  每一列的标题
     * @return
     */
    public static HSSFWorkbook writeExcel(List objects, String[] fields, String[] titles) {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        HSSFRow row = sheet.createRow(0);
        HSSFCellStyle style = wb.createCellStyle();
        style.setAlignment(HorizontalAlignment.GENERAL);
        for (int i = 0; i < titles.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellValue(titles[i]);
            cell.setCellStyle(style);
//            sheet.autoSizeColumn(i);
        }
        for (int i = 0; i < objects.size(); i++) {
            row = sheet.createRow(i + 1);
            Object obj = objects.get(i);
            Class classType = obj.getClass();
            int cellIndex = 0;
            for (String field : fields) {
                String firstLetter = field.substring(0, 1).toUpperCase();
                String getMethodName = "get" + firstLetter + field.substring(1);
                Method getMethod = null;
                Object value = null;

                try {
                    getMethod = classType.getMethod(getMethodName, new Class[]{});
                    value = getMethod.invoke(obj, new Object[]{});
                } catch (NoSuchMethodException e) {
                    log.error("there is no such method: {}", getMethod.toString());
                    e.printStackTrace();
                } catch (IllegalAccessException e) {
                    log.error("this is an illegal access Method return value: {}", getMethod.toString());
                    e.printStackTrace();
                } catch (InvocationTargetException e) {
                    log.error("this is an InvocationTargetException: {}", getMethod.toString());
                    e.printStackTrace();
                }
                row.createCell(cellIndex).setCellValue(value == null ? "" : value.toString());
                cellIndex++;
            }
        }
        return wb;
    }

    /**
     * String集合写sheet
     * 写入titles表头和list每一行数据的集合
     * @param sheet
     * @param titles
     * @param list
     * @return
     */
    public static Sheet writeTableOnSheet(Sheet sheet, List<String> titles, List<String[]> list) {

//        HSSFSheet sheet1 = workbook.createSheet("直播时长汇总");
//        List<String> titles = Stream.of(new String[]{"用户ID", "昵称", "用户IP", "用户IP区域", "观看总时长", "PC端总时长", "移动端总时长"}).collect(Collectors.toList());
//        writeTableOnSheet(sheet1, titles, rowList);
        if (sheet != null) {
            int rownumInit = sheet.getLastRowNum() == 0 ? 0 : sheet.getLastRowNum() + 1;
            Row row = sheet.createRow(rownumInit);
            //1.表头
            for (int i = 0; i < titles.size(); i++) {
                Cell cell = row.createCell(i);
                cell.setCellValue(titles.get(i));
            }
            //2.表数据
            if (list != null) {
                for (int i = 0; i < list.size(); i++) {
                    sheet = writeRowOnSheet(sheet, list.get(i));
                }
            }
        }
        return sheet;
    }

    /**
     * String数组写row
     * sheet中写一行strs数组
     * @param sheet
     * @param strs
     * @return
     */
    public static Sheet writeRowOnSheet(Sheet sheet, String[] strs) {
        Row hssfRow = sheet.createRow(sheet.getLastRowNum() + 1);
        int i = 0;
        for (String str : strs) {
            Cell cell = hssfRow.createCell(i++);
            cell.setCellValue(str);
        }
        return sheet;
    }

    //----------------------------------------private method---------------------------------

    /**
     * 获取相应版本的workbook
     * @param inStr
     * @param fileName
     * @return
     */
    private static Workbook getWorkbook(InputStream inStr, String fileName) {
        Workbook wb = null;
        String fileType = fileName.substring(fileName.lastIndexOf("."));
        try {
            if (EXCEL2003.equals(fileType)) {
                wb = new HSSFWorkbook(inStr);
            } else if (EXCEL2007.equals(fileType)) {
                wb = new XSSFWorkbook(inStr);
            }
        } catch (Exception e) {
            log.error("解析的文件格式有误！");
        }
        return wb;
    }

    /**
     * 实体类的字段赋值
     * @param t 实体类
     * @param value 字段值
     * @param field 字段
     * @throws Exception
     */
    public static <T> void fillField(T t, String value, Field field) throws Exception {
        Class<?> type = field.getType();
        if (type == null || type == void.class || StringUtils.isBlank(value)) {
            return;
        }
        if (type == Object.class) {
            field.set(t, value);
            //数字
        } else if (type.getSuperclass() == null || type.getSuperclass() == Number.class) {
            if (type == int.class || type == Integer.class) {
                field.set(t, NumberUtils.toInt(value));
            } else if (type == long.class || type == Long.class) {
                field.set(t, NumberUtils.toLong(value));
            } else if (type == byte.class || type == Byte.class) {
                field.set(t, NumberUtils.toByte(value));
            } else if (type == short.class || type == Short.class) {
                field.set(t, NumberUtils.toShort(value));
            } else if (type == double.class || type == Double.class) {
                field.set(t, NumberUtils.toDouble(value));
            } else if (type == float.class || type == Float.class) {
                field.set(t, NumberUtils.toFloat(value));
            } else if (type == char.class || type == Character.class) {
                field.set(t, CharUtils.toChar(value));
            } else if (type == boolean.class) {
                field.set(t, BooleanUtils.toBoolean(value));
            } else if (type == BigDecimal.class) {
                field.set(t, new BigDecimal(value));
            }
        } else if (type == Boolean.class) {
            field.set(t, BooleanUtils.toBoolean(value));
        } else if (type == Date.class) {
            //
            field.set(t, value);
        } else if (type == String.class) {
            field.set(t, value);
        } else {
            Constructor<?> constructor = type.getConstructor(String.class);
            field.set(t, constructor.newInstance(value));
        }
    }

    /**
     * 浏览器下载excel
     * @param fileName 文件名
     * @param workbook
     */
    private static void downloadExcel(String fileName, Workbook workbook, HttpServletResponse response){
        try {
//            HttpServletResponse response = ((ServletRequestAttributes) RequestContextHolder.getRequestAttributes()).getResponse();
            response.setContentType(MediaType.APPLICATION_OCTET_STREAM_VALUE);
            response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));
            response.flushBuffer();
            workbook.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }


    /**
     * 输出文件 path和file只需传一个
     * @param path 生成excel路径
     * @param file 文件
     * @param workbook
     */
    private static boolean outFile(String path, File file, Workbook workbook) {
        OutputStream out = null;
        //按path创建文件
        if(StringUtils.isNotBlank(path)) {
            file = new File(path);
            if (file.exists()) {
                file.delete();
            }
        }
        try {
            out = new FileOutputStream(file);
            workbook.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (out != null) {
                try {
                    out.close();
                } catch (Exception ex) {
                    ex.printStackTrace();
                }
            }
        }
        return file.exists();
    }

    //----------------------------------------类型转换----------------------------------------

    private static Integer stringToInt(String str) {
        if (str != null && str.contains(".")) {
            str = str.substring(0, str.indexOf("."));
        }
        return StringUtils.isEmpty(str) ? null : NumberUtils.toInt(str);
    }

    private static Double stringToDouble(String str) {
        return StringUtils.isEmpty(str) ? null : NumberUtils.toDouble(str);
    }

    private static Long stringToLong(String str) {
        if (str != null && str.contains(".")) {
            str = str.substring(0, str.indexOf("."));
        }
        return StringUtils.isEmpty(str) ? null : NumberUtils.toLong(str);
    }

    /**
     * 将列名和list转换成map
     * @param fristRow
     * @param values
     * @return
     */
    private static Map<String, String> rowToMap(Row fristRow, List<String> values) {
        Map<String, String> rowMap = new LinkedHashMap();
        for (int i = 0; i < values.size(); i++) {
            String name = fristRow.getCell(i).toString();
            String value = values.get(i);
            if (StringUtils.isBlank(value) || StringUtils.isBlank(name)) {
                continue;
            }
            rowMap.put(name.trim(), value.trim());
        }
        return rowMap;
    }


}
