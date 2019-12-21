package com.bzc.example.demo.utils;

import com.bzc.example.demo.constants.GlobalConstants;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;

import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.text.SimpleDateFormat;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

@Slf4j
public class ExportExcelUtil<T> {

    /**
     * 这是一个通用的方法，利用了JAVA的反射机制，可以将放置在JAVA集合中并且符号一定条件的数据以EXCEL 的形式输出到指定IO设备上
     *
     * @param title
     *            表格标题名
     * @param headers
     *            表格属性列名数组
     * @param dataset
     *            需要显示的数据集合,集合中一定要放置符合javabean风格的类的对象。此方法支持的
     *            javabean属性的数据类型有基本数据类型及String,Date,byte[](图片数据)
     * @param out
     *            与输出设备关联的流对象，可以将EXCEL文档导出到本地文件或者网络中
     * @param pattern
     *            如果有时间数据，设定输出格式。默认为"yyy-MM-dd"
     */
    @SuppressWarnings("unchecked")
    public void exportForExcel(String title, String[] headers, Collection<T> dataset, OutputStream out,
                               String pattern) {
        // 声明一个工作薄
        HSSFWorkbook workbook = new HSSFWorkbook();
        // 生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为18个字节
        sheet.setDefaultColumnWidth(18);
        // 生成一个样式
        HSSFCellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor((short) 9);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        // 生成一个字体
        HSSFFont font = workbook.createFont();
        font.setColor(Font.COLOR_NORMAL);
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        // 把字体应用到当前的样式
        style.setFont(font);
        // 生成并设置另一个样式
        HSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor((short) 9);
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        // 生成另一个字体
        HSSFFont font2 = workbook.createFont();
        font2.setBold(true);
        // 把字体应用到当前的样式
        style2.setFont(font2);
        // 声明一个画图的顶级管理器
        HSSFPatriarch patriarch = sheet.createDrawingPatriarch();
        // 定义注释的大小和位置,详见文档
        // HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0, 0, 0,
        // 0, (short) 4, 2, (short) 6, 5));
        // 设置注释内容
        // comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
        // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
        // comment.setAuthor("leno");
        // 产生表格标题行
        HSSFRow row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            HSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        // 遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        HSSFFont font3 = workbook.createFont();
        font3.setColor((short) 8);
        while (it.hasNext()) {
            index++;
            row = sheet.createRow(index);
            T t = it.next();
            // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
            Field[] fields = t.getClass().getDeclaredFields();
            for (int i = 0; i < headers.length; i++) {
                HSSFCell cell = row.createCell(i);
                cell.setCellStyle(style2);
                Field field = fields[i];
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                try {
                    @SuppressWarnings("rawtypes")
                    Class tCls = t.getClass();
                    Method getMethod = tCls.getMethod(getMethodName);
                    Object value = getMethod.invoke(t);
                    // 判断值的类型后进行强制类型转换
                    String textValue = null;
                    if (value instanceof Date) {
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        textValue = sdf.format(date);
                    } else if (value instanceof byte[]) {
                        // 有图片时，设置行高为60px;
                        row.setHeightInPoints(60);
                        // 设置图片所在列宽度为80px,注意这里单位的一个换算
                        sheet.setColumnWidth(i, (short) (35.7 * 80));
                        // sheet.autoSizeColumn(i);
                        byte[] bsValue = (byte[]) value;
                        HSSFClientAnchor anchor = new HSSFClientAnchor(
                                0, 0, 1023, 255, (short) 6, index, (short) 6, index);
                        anchor.setAnchorType(ClientAnchor.AnchorType.byId(2));
                        patriarch.createPicture(anchor, workbook.addPicture(bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                    } else {
                        // 其它数据类型都当作字符串简单处理
                        textValue = String.valueOf(value);
                    }
                    // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                    if (null == textValue || "null".equals(textValue)) {
                        textValue = "";
                    }
                    Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                    Matcher matcher = p.matcher(textValue);
                    if (matcher.matches()) {
                        // 是数字当作double处理
                        cell.setCellValue(Double.parseDouble(textValue));
                    } else {
                        HSSFRichTextString richString = new HSSFRichTextString(textValue);
                        // font3.setColor(HSSFColor.BLUE.index);
                        richString.applyFont(font3);
                        cell.setCellValue(richString);
                    }
                } catch (SecurityException | NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
        try {
            workbook.write(out);
            out.flush();
            //out.close();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }



    // 导出时分多个sheet页
    public void exportForExcelMoreSheets(String title, String[] headers, Collection<? extends Collection<T>> dataSet,
                                         OutputStream out, String pattern) {
        HSSFWorkbook workbook = new HSSFWorkbook(); // 声明一个工作薄
        HSSFSheet[] sheets = new HSSFSheet[dataSet.size()];
        HSSFCellStyle style = workbook.createCellStyle(); // 生成一个样式
        style.setFillForegroundColor((short) 9);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        HSSFFont font = workbook.createFont(); // 生成一个字体
        font.setColor(Font.COLOR_NORMAL);
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        style.setFont(font); // 把字体应用到当前的样式
        HSSFCellStyle style2 = workbook.createCellStyle(); // 生成并设置另一个样式
        style2.setFillForegroundColor((short) 9);
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        HSSFFont font2 = workbook.createFont(); // 生成另一个字体
        font2.setBold(true);
        style2.setFont(font2); // 把字体应用到当前的样式
        HSSFFont font3 = workbook.createFont();
        font3.setColor((short) 8);
        // 遍历数据
        int pos = 0;
//        Iterator<? extends Collection<T>> items = dataSet.iterator();
        for (Collection<T> item : dataSet) {
            sheets[pos] = workbook.createSheet(title + "_" + pos); // 生成一个表格
            sheets[pos].setDefaultColumnWidth(18);
            HSSFPatriarch patriarch = sheets[pos].createDrawingPatriarch(); // 声明一个画图的顶级管理器
            HSSFRow row = sheets[pos].createRow(0); // 产生表格标题行
            for (int loc = 0; loc < headers.length; loc++) {
                HSSFCell cell = row.createCell(loc);
                cell.setCellStyle(style);
                cell.setCellValue(new HSSFRichTextString(headers[loc]));
            }
            Iterator<T> it = item.iterator();
            int index = 0;
            while (it.hasNext()) {
                index++;
                row = sheets[pos].createRow(index);
                T t = it.next();
                // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
                Field[] fields = t.getClass().getDeclaredFields();
                for (int line = 0; line < headers.length; line++) {
                    HSSFCell cell = row.createCell(line);
                    cell.setCellStyle(style2);
                    Field field = fields[line];
                    String fieldName = field.getName();
                    String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                    try {
                        Class clazz = t.getClass();
                        Method getMethod = clazz.getMethod(getMethodName);
                        Object value = getMethod.invoke(t);
                        String textValue = null; // 判断值的类型后进行强制类型转换
                        if (value instanceof Date) {
                            Date date = (Date) value;
                            SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                            textValue = sdf.format(date);
                        } else if (value instanceof byte[]) {
                            row.setHeightInPoints(60); // 有图片时，设置行高为60px;
                            sheets[pos].setColumnWidth(pos, (short) (35.7 * 80)); // 设置图片所在列宽度为80px,注意这里单位的一个换算
                            // sheet.autoSizeColumn(i);
                            byte[] bsValue = (byte[]) value;
                            HSSFClientAnchor anchor = new HSSFClientAnchor(
                                    0, 0, 1023, 255, (short) 6, index, (short) 6, index);
                            anchor.setAnchorType(ClientAnchor.AnchorType.byId(2));
                            patriarch.createPicture(anchor, workbook.addPicture(bsValue, HSSFWorkbook.PICTURE_TYPE_JPEG));
                        } else {
                            textValue = String.valueOf(value); // 其它数据类型都当作字符串简单处理
                        }
                        if (null == textValue || "null".equals(textValue)) {
                            textValue = ""; // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                        }
                        Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                        Matcher matcher = p.matcher(textValue);
                        if (matcher.matches()) {
                            cell.setCellValue(Double.parseDouble(textValue)); // 是数字当作double处理
                        } else {
                            HSSFRichTextString richString = new HSSFRichTextString(textValue);
                            richString.applyFont(font3);
                            cell.setCellValue(richString);
                        }
                    } catch (NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                        log.error(e.getMessage(), e);
                    }
                }
            }
            pos++;
        }
        try {
            workbook.write(out);
            out.flush();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }

    /**
     * 将一个list均分成n个list
     * @param source 待分割的集合
     * @return 结果
     */
    public List<List<T>> averageAssign(List<T> source) {
        List<List<T>> result = new ArrayList<>();
        int number = GlobalConstants.RECORDS_MAX_NUMBER;
        int n = (int) Math.ceil(source.size() * 1.0 / number);
        List<T> value;
        for (int index = 0; index < n; index++) {
            if (index == n - 1) {
                value = source.subList(index * number, source.size());
            } else {
                value = source.subList(index * number, (index + 1) * number);
            }
            result.add(value);
        }
        return result;
    }


    public void exportForExcel2007(String title, String[] headers, Collection<T> dataset, OutputStream out,
                                   String pattern) {
        // 声明一个工作薄
        XSSFWorkbook workbook = new XSSFWorkbook();
        // 生成一个表格
        XSSFSheet sheet = workbook.createSheet(title);
        // 设置表格默认列宽度为18个字节
        sheet.setDefaultColumnWidth(18);

        XSSFDataFormat format = workbook.createDataFormat();

        // 生成一个样式
        XSSFCellStyle style = workbook.createCellStyle();
        // 设置这些样式
        style.setFillForegroundColor((short) 9);
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);
        style.setBorderTop(BorderStyle.THIN);
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setDataFormat(format.getFormat("@"));
        // 生成一个字体
        XSSFFont font = workbook.createFont();
        font.setColor(Font.COLOR_NORMAL);
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        // 把字体应用到当前的样式
        style.setFont(font);


        // 生成并设置另一个样式
        XSSFCellStyle style2 = workbook.createCellStyle();
        style2.setFillForegroundColor((short) 9);
        style2.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        style2.setBorderBottom(BorderStyle.THIN);
        style2.setBorderLeft(BorderStyle.THIN);
        style2.setBorderRight(BorderStyle.THIN);
        style2.setBorderTop(BorderStyle.THIN);
        style2.setAlignment(HorizontalAlignment.CENTER);
        style2.setVerticalAlignment(VerticalAlignment.CENTER);
        style2.setDataFormat(format.getFormat("@"));
        // 生成另一个字体
        XSSFFont font2 = workbook.createFont();
        font2.setBold(true);
        // 把字体应用到当前的样式
        style2.setFont(font2);
        // 声明一个画图的顶级管理器
        XSSFDrawing patriarch = sheet.createDrawingPatriarch();
        // 定义注释的大小和位置,详见文档
        // HSSFComment comment = patriarch.createComment(new HSSFClientAnchor(0, 0, 0,
        // 0, (short) 4, 2, (short) 6, 5));
        // 设置注释内容
        // comment.setString(new HSSFRichTextString("可以在POI中添加注释！"));
        // 设置注释作者，当鼠标移动到单元格上是可以在状态栏中看到该内容.
        // comment.setAuthor("leno");
        // 产生表格标题行
        XSSFRow row = sheet.createRow(0);

        for (int i = 0; i < headers.length; i++) {
            XSSFCell cell = row.createCell(i);
            cell.setCellStyle(style);
            XSSFRichTextString text = new XSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }
        // 遍历集合数据，产生数据行
        Iterator<T> it = dataset.iterator();
        int index = 0;
        XSSFFont font3 = workbook.createFont();
        font3.setColor((short) 8);

        while (it.hasNext()) {
            index++;
            row = sheet.createRow(index);
            T t = it.next();
            // 利用反射，根据javabean属性的先后顺序，动态调用getXxx()方法得到属性值
            Field[] fields = t.getClass().getDeclaredFields();

            for (int i = 0; i < headers.length; i++) {
                XSSFCell cell = row.createCell(i);

                cell.setCellType(XSSFCell.CELL_TYPE_STRING);
                cell.setCellStyle(style2);
                Field field = fields[i];
                String fieldName = field.getName();
                String getMethodName = "get" + fieldName.substring(0, 1).toUpperCase() + fieldName.substring(1);
                try {
                    @SuppressWarnings("rawtypes")
                    Class tCls = t.getClass();
                    Method getMethod = tCls.getMethod(getMethodName);
                    Object value = getMethod.invoke(t);
                    // 判断值的类型后进行强制类型转换
                    String textValue = null;
                    if (value instanceof Date) {
                        Date date = (Date) value;
                        SimpleDateFormat sdf = new SimpleDateFormat(pattern);
                        textValue = sdf.format(date);
                    } else if (value instanceof byte[]) {
                        // 有图片时，设置行高为60px;
                        row.setHeightInPoints(60);
                        // 设置图片所在列宽度为80px,注意这里单位的一个换算
                        sheet.setColumnWidth(i, (short) (35.7 * 80));
                        // sheet.autoSizeColumn(i);
                        byte[] bsValue = (byte[]) value;
                        XSSFClientAnchor anchor = new XSSFClientAnchor(
                                0, 0, 1023, 255, (short) 6, index, (short) 6, index);
                        anchor.setAnchorType(ClientAnchor.AnchorType.byId(2));
                        patriarch.createPicture(anchor, workbook.addPicture(bsValue, XSSFWorkbook.PICTURE_TYPE_JPEG));
                    } else {
                        // 其它数据类型都当作字符串简单处理
                        textValue = String.valueOf(value);
                    }
                    // 如果不是图片数据，就利用正则表达式判断textValue是否全部由数字组成
                    if (null == textValue || "null".equals(textValue)) {
                        textValue = "";
                    }
                    Pattern p = Pattern.compile("^//d+(//.//d+)?$");
                    Matcher matcher = p.matcher(textValue);
                    if (matcher.matches()) {
                        // 是数字当作double处理
                        cell.setCellValue(Double.parseDouble(textValue));
                    } else {
                        XSSFRichTextString richString = new XSSFRichTextString(textValue);
                        // font3.setColor(HSSFColor.BLUE.index);
                        richString.applyFont(font3);
                        cell.setCellValue(richString);
                    }
                } catch (SecurityException | NoSuchMethodException | IllegalAccessException | InvocationTargetException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
        try {
            workbook.write(out);
            out.flush();
            //out.close();
        } catch (IOException e) {
            log.error(e.getMessage(), e);
        }
    }
}