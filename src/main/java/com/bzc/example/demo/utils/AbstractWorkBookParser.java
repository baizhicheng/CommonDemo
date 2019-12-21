package com.bzc.example.demo.utils;

import com.bzc.example.demo.annotation.Excel;
import com.bzc.example.demo.annotation.ExcelFied;
import lombok.extern.slf4j.Slf4j;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.NumberToTextConverter;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.*;

@Slf4j
@Component
public abstract class AbstractWorkBookParser {

    @Autowired
    private JdbcTemplate jdbcTemplate;


    public Map<String,Object> parseExcel(Workbook workbook, String fileName, Class<?> clazz) {
        // 先计算所有公式
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        //sheet读取页集合
        String[] sheets={};

        //起始行数
        int skip=0;

        Annotation[] annotations = clazz.getAnnotations();
        Excel _excel;

        for(Annotation annotation:annotations){
            if(annotation == null){ continue; }

            if (annotation instanceof Excel) {
                _excel= (Excel)annotation;
                //读取注解内容
                skip = _excel.skip();
                sheets = _excel.sheets();

                break;
            }
        }

        CellStyle styleDate = createCellStyle(workbook);

        //获取总的sheet页数量
        int sheetCount = workbook.getNumberOfSheets();  //Sheet的数量

        //判断是否解析所有sheets
        Map<String,Object> map = null;
        int sheetIndex = 0;

        if(Arrays.asList(sheets).contains(Excel.ALL)){
            //遍历每个Sheet
            for (int s = 0; s < sheetCount; s++) {
                Sheet sheet = workbook.getSheetAt(s);
                map = this.readSheetRow(sheet,skip,clazz,styleDate,evaluator);
                if (null != map) map.put(ExcelUtils.SHEET_TITLES, this.getSheetTitle(sheet, skip));
            }
        }else{
            //遍历指定Sheets
            for (int s = 0; s < sheets.length; s++) {
                //如果sheet不存在继续寻找下一个sheet
                if(Integer.valueOf(sheets[s])>sheetCount){
                    continue;
                }

                Sheet sheet = workbook.getSheetAt(Integer.valueOf(sheets[s]));
                map = this.readSheetRow(sheet,skip,clazz,styleDate,evaluator);

                if (null != map) map.put(ExcelUtils.SHEET_TITLES, this.getSheetTitle(sheet, skip));

                sheetIndex = Integer.valueOf(sheets[s]);
            }
        }

        if (map != null) {
            map.put(ExcelUtils.FILE_NAME,fileName);
            map.put(ExcelUtils.SAVE_FILE,workbook);
            map.put(ExcelUtils.SHEET_INDEX,sheetIndex);
        }

        return map;
    }


    public Map<String,Object> parseExcel(Workbook workbook, int sheetIndex, String fileName, Class<?> clazz) {
        Sheet sheet = workbook.getSheetAt(sheetIndex);

        if (sheet == null){
            return null;
        }

        // 先计算所有公式
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();

        //起始行数
        int skip=0;

        Annotation[] annotations = clazz.getAnnotations();
        Excel _excel;

        for(Annotation annotation:annotations){
            if(annotation == null){ continue; }

            if (annotation instanceof Excel) {
                _excel= (Excel)annotation;
                //读取注解内容
                skip = _excel.skip();

                break;
            }
        }

        CellStyle styleDate = createCellStyle(workbook);

        //判断是否解析所有sheets
        Map<String,Object> map = this.readSheetRow(sheet,skip,clazz,styleDate,evaluator);

        if (null != map) map.put(ExcelUtils.SHEET_TITLES, this.getSheetTitle(sheet, skip));


        if (map != null) {
            map.put(ExcelUtils.FILE_NAME,fileName);
            map.put(ExcelUtils.SAVE_FILE,workbook);
            map.put(ExcelUtils.SHEET_INDEX,sheetIndex);
        }

        return map;
    }



    private String[] getSheetTitle(Sheet sheet) {
        Row titleRow = sheet.getRow(0);
        int columns = titleRow.getPhysicalNumberOfCells();
        String[] titles = new String[columns];

        for (int col = 0; col < columns; col++) {
            Cell cell = titleRow.getCell(col);
            titles[col] = cell.getStringCellValue();
        }

        return titles;
    }

    private String[] getSheetTitle(Sheet sheet, int index) {
        Row titleRow = sheet.getRow(index);
        int columns = titleRow.getPhysicalNumberOfCells();
        String[] titles = new String[columns];

        for (int col = 0; col < columns; col++) {
            Cell cell = titleRow.getCell(col);
            titles[col] = cell.getStringCellValue();
        }

        return titles;
    }

    /**
     *  增加错误备注
     * @param workbook 文本
     * @param map 解析文本的值
     * @param errorMap 错误信息
     */
    @SuppressWarnings("unchecked")
    public void addErrorComment(Workbook workbook, Map<String, Object> map, Map<Integer, Object> errorMap) {
        CellStyle styleDate = createCellStyle(workbook);

        if (errorMap == null || errorMap.size() == 0)
            return;

        Map<Integer,Integer> rowIndexMap = (Map<Integer, Integer>) map.get(ExcelUtils.ROW_INDEX);
        Map<String,Integer> excelFieldMap = (Map<String, Integer>) map.get(ExcelUtils.EXCEL_FIELD_INDEX);
        Iterator<Integer> iterator = errorMap.keySet().iterator();
        Sheet sheet = workbook.getSheetAt(Integer.valueOf(map.get(ExcelUtils.SHEET_INDEX).toString()));


        while (iterator.hasNext()) {
            int index = iterator.next();
            int r = rowIndexMap.get(index);
            Row row = sheet.getRow(r);
            Map<String,String> fieldMsg = (Map<String, String>) errorMap.get(index);
            Iterator<Map.Entry<String, String>> fieldIterator = fieldMsg.entrySet().iterator();

            while (fieldIterator.hasNext()) {
                Map.Entry<String,String> entry = fieldIterator.next();
                String fieldName = entry.getKey();
                String errorMsg = entry.getValue();
                int cellIndex = excelFieldMap.get(fieldName);
                Cell cell = row.getCell(cellIndex);

                if(null == cell)
                    cell=row.createCell(cellIndex);

                if(null == cell.getCellComment()){//不存在注释则创建
                    Comment tcomment = null;

                    try{
//                        tcomment = setComment(sheet,errorMsg);
//                        tcomment = setCommentByAnchorSingle(sheet, errorMsg, cellIndex, r);
                        tcomment = setCommentByAnchor(sheet, errorMsg, cellIndex, r);
                    }catch(Exception e){
                        log.error(e.getMessage());
                    }finally {
                        if(null != tcomment)
                            cell.setCellComment(tcomment);
                    }
                }else
                    cell.getCellComment().setString(new XSSFRichTextString(errorMsg));


                styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                cell.setCellStyle(styleDate);
            }
        }
    }

    abstract CellStyle createCellStyle(Workbook workbook);


    /**
     *  读取excel中的每行
     * @param sheet 数据sheet页
     * @param skip 标题起始行
     * @param clazz 模板类
     * @param styleDate 样式
     * @return 回复map
     */
    private Map<String,Object> readSheetRow(Sheet sheet, int skip, Class<?> clazz,
                                            CellStyle styleDate, FormulaEvaluator evaluator) {
        Map<String,Integer> excelFieldIndex = new HashMap<>();
        Field[] fields = clazz.getDeclaredFields();
        Map<ExcelFied, Field> textToKey = new HashMap<>();
        ExcelFied _excelField;

        for (Field field : fields) {
            _excelField = field.getAnnotation(ExcelFied.class);

            if (_excelField == null  ) { continue; }

            textToKey.put(_excelField, field);
            excelFieldIndex.put(field.getName(),_excelField.index());
        }

        List<Object> result = new ArrayList<>();
        Map<Integer,Integer> listMap = new HashMap<>();

//        int rowCount = sheet.getPhysicalNumberOfRows(); //获取总行数(有效行数)
        int rowCount = sheet.getLastRowNum(); //获取总行数（实际行数）

        if(rowCount < skip + 1){ return null; }
        // 非法数据计数
        int invalidCount = 0;
        // Excel数据总数
        int totalCount = 0;
        //遍历每一行
        Row titlerow = sheet.getRow(skip);

        if (titlerow == null){ return null; }

        int coloumnCountNumber = titlerow.getPhysicalNumberOfCells();//获取总列数(有效列数)

        for (int r = skip + 1; r <= rowCount; r++) {
            Row row = sheet.getRow(r);

            if(row == null){
                break;
            }

            //int cellCount = row.getPhysicalNumberOfCells(); //获取总列数
            int cellCount = coloumnCountNumber; //获取总列数

            // 检查是否是最后一行
            int totalNull = 0;

            for (int c = 0; c < cellCount;c ++) {
                Cell cell = row.getCell(c);
                if (cell == null || StringUtils.isEmpty(cell.toString())) {
                    totalNull ++ ;
                }
            }

            if (totalNull == cellCount) { break; }

            totalCount ++;
            Object object = null;

            try {
                object= clazz.newInstance();
            } catch (Exception e) {
                e.printStackTrace();
            }

            // 每行单元格校验失败的次数
            int failCount = 0;

            //遍历给指定属性赋值(只遍历注解指定的列)
            for(ExcelFied key : textToKey.keySet()){
                if(key.index()>cellCount) {
                    continue;
                }

                Cell cell = row.getCell(key.index());
                Object cellValue = null;

                //获取合并单元格数据的值
                if(isMergedRegion(sheet,r,key.index())) {
                    int sheetMergeCount = sheet.getNumMergedRegions();

                    for (int i = 0; i < sheetMergeCount; i++) {
                        CellRangeAddress ca = sheet.getMergedRegion(i);

                        int firstColumn = ca.getFirstColumn();
                        int lastColumn = ca.getLastColumn();
                        int firstRow = ca.getFirstRow();
                        int lastRow = ca.getLastRow();

                        if (r >= firstRow && r <= lastRow) {
                            if (key.index() >= firstColumn && key.index() <= lastColumn) {
                                Row fRow = sheet.getRow(firstRow);
                                cell = fRow.getCell(firstColumn);
                            }
                        }
                    }
                }

                try {
                    if (cell == null) {
                        cell = row.createCell(key.index());
                    }

                    if (cell.getColumnIndex() == 0){
                        System.out.println(0);
                    }

                    cellValue = this.getCellValue(sheet,cell,key, styleDate,evaluator);
                }catch (Exception e) {
                    failCount ++;
                }

                if (failCount == 0) {
                    Field field=  textToKey.get(key);

                    try {
                        field.setAccessible(true);
                        BeanUtils.setProperty(object, field.getName(), cellValue);
                    } catch (Exception e) {
                        log.error(e.getMessage(), e);
                    }
                } else {
                    object = null;
                }
            }

            if (object != null) {
                // 合法数据加入list
                result.add(object);
                listMap.put(result.size() - 1,r);
            } else {
                // 非法数据计数+1
                invalidCount++;
            }
        }

        Map<String,Object> resultMap = new HashMap<>();

        resultMap.put(ExcelUtils.TOTAL_COUNT,totalCount);
        resultMap.put(ExcelUtils.INVALID_COUNT,invalidCount);
        resultMap.put(ExcelUtils.VALID_LIST,result);
        resultMap.put(ExcelUtils.ROW_INDEX,listMap);
        resultMap.put(ExcelUtils.EXCEL_FIELD_INDEX,excelFieldIndex);

        return resultMap;
    }

    /**
     *  获取单元格值
     * @param sheet
     * @param cell
     * @param key
     * @param styleDate
     * @return
     */
    private  Object getCellValue(Sheet sheet, Cell cell, ExcelFied key, CellStyle styleDate, FormulaEvaluator evaluator){
        Object cellValue;
        SimpleDateFormat fmt = new SimpleDateFormat("yyyy-MM-dd");
        SimpleDateFormat timeFmt = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        int cellType = cell.getCellType();

        switch (cellType) {
            case Cell.CELL_TYPE_STRING: //文本
                cellValue = cell.getStringCellValue().trim();
                break;
            case Cell.CELL_TYPE_NUMERIC: //数字、日期
                if (DateUtil.isCellDateFormatted(cell)) {
//                    // 第一行第十一列标题为“采集时间”时做特殊处理，需要读取完整日期和时间
//                    Cell timeTitleCell = sheet.getRow(0).getCell(11);
//
//                    if (timeTitleCell.getCellType() == Cell.CELL_TYPE_STRING
//                    && timeTitleCell.getStringCellValue().equals("采集时间"))
//                        cellValue = timeFmt.format(cell.getDateCellValue()); //日期时间型
//                    else
//                        cellValue = fmt.format(cell.getDateCellValue()); //日期型

                    cellValue = timeFmt.format(cell.getDateCellValue()); //日期时间型
                } else {
                    // 解决问题：1，科学计数法(如2.6E+10)，2，超长小数小数位不一致（如1091.19649281798读取出1091.1964928179796），3，整型变小数（如0读取出0.0）
                    cellValue = NumberToTextConverter.toText(cell.getNumericCellValue());
                }
                break;
            case Cell.CELL_TYPE_BOOLEAN: //布尔型
                cellValue = cell.getBooleanCellValue();
                break;
            case Cell.CELL_TYPE_BLANK: //空白
                cellValue = cell.getStringCellValue();
                break;
            case Cell.CELL_TYPE_FORMULA: //公式
                CellValue cellValue1 =  evaluator.evaluate(cell);
                cellValue = "0";
                switch (cellValue1.getCellType()) {
                    case Cell.CELL_TYPE_NUMERIC:
                        BigDecimal bg = new BigDecimal(cellValue1.getNumberValue());
                        cellValue =  bg.setScale(2, BigDecimal.ROUND_HALF_UP).doubleValue();
                        break;
                    case Cell.CELL_TYPE_STRING:
                        cellValue =  cellValue1.getStringValue();
                        break;
                    case Cell.CELL_TYPE_BOOLEAN:
                        cellValue =  cellValue1.getBooleanValue();
                        break;
                    case Cell.CELL_TYPE_ERROR:
                        if (null != cell.getCellComment()) {
                            cell.removeCellComment();// 删除批注
                        }
                        cell.setCellComment(setComment(sheet,"错误计算"));
                        styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                        cell.setCellStyle(styleDate);
                        throw new RuntimeException("错误计算");
                }
                break;
            case Cell.CELL_TYPE_ERROR: //错误
                cellValue = "0";
                break;
            default:
                cellValue = "0";
        }
        if (key.isTrim() ) {
            cellValue = cellValue.toString().replaceAll(" ","");
        }
        while (true) {
            if (key.notNull()) { // 不能为空
                if (StringUtils.isNullOrEmptyString(cellValue.toString())) {
                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }
                    cell.setCellComment(setComment(sheet,"不能为空"));
                    styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    cell.setCellStyle(styleDate);
                    throw new RuntimeException("不能为空");
                }
            }
            // 如果单元格为空，不做后面判断
            if (StringUtils.isEmpty(cellValue.toString())) {
                break;
            }
            if (key.length() > 0) { // 判断长度
                if (cellValue.toString().length() > key.length()) {
                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }
                    cell.setCellComment(setComment(sheet,"长度最大为" + key.length()));
                    styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    cell.setCellStyle(styleDate);
                    throw new RuntimeException("超过最大长度");
                }
            }
            switch (key.type()) {
                case STRING:
                    cellValue = cellValue.toString();
                    break;
                case NUMBER:
                    try {
                        if (!isIntegerForDouble( Double.parseDouble(cellValue.toString()))) {
                            if (null != cell.getCellComment()) {
                                cell.removeCellComment();// 删除批注
                            }
                            cell.setCellComment(setComment(sheet,"不是整数型" ));
                            styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                            cell.setCellStyle(styleDate);
                            throw new RuntimeException("不是整数型");
                        }
                    } catch (Exception e) {
                        if (null != cell.getCellComment()) {
                            cell.removeCellComment();// 删除批注
                        }
                        cell.setCellComment(setComment(sheet,"不是整数型" ));
                        styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                        cell.setCellStyle(styleDate);
                        throw new RuntimeException("不是整数型");
                    }
                    break;
                case DOUBLE:
                    try {
                        Double.parseDouble(cellValue.toString());
                    } catch (Exception e) {
                        if (null != cell.getCellComment()) {
                            cell.removeCellComment();// 删除批注
                        }
                        cell.setCellComment(setComment(sheet,"不是浮点数型" ));
                        styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                        cell.setCellStyle(styleDate);
                        throw new RuntimeException("不是浮点数型");
                    }
                    break;
                case BOOLEAN:
                    try {
                        Boolean.parseBoolean(cellValue.toString());
                    } catch (Exception e) {
                        if (null != cell.getCellComment()) {
                            cell.removeCellComment();// 删除批注
                        }
                        cell.setCellComment(setComment(sheet,"不是布尔型" ));
                        styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                        cell.setCellStyle(styleDate);
                        throw new RuntimeException("不是布尔型");
                    }
                    break;
                case DATE:
                    break;
                case DATETIME:
                    break;
            }
            // 最小值
            if (!StringUtils.isEmpty(key.minValue()) && !StringUtils.isEmpty(cellValue.toString())) {
                // 如果单元的值小于设置的最小值
                if (!compareValue(cellValue.toString(),key.minValue())) {
                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }
                    cell.setCellComment(setComment(sheet,"最小值是" + key.minValue() ));
                    styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    cell.setCellStyle(styleDate);
                    throw new RuntimeException("小于设置最小值");
                }
            }
            // 最大值
            if (!StringUtils.isEmpty(key.maxValue()) && !StringUtils.isEmpty(cellValue.toString())) {
                // 如果单元的值大于设置的最小值
                if (compareValue(cellValue.toString(),key.maxValue())) {
                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }
                    cell.setCellComment(setComment(sheet,"最大值是" + key.maxValue() ));
                    styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    cell.setCellStyle(styleDate);
                    throw new RuntimeException("超过设置最大值");
                }
            }
            //判断是否需要直接查询
            if(!StringUtils.isEmpty(key.sqlStr())){
//                String sqlStr = key.sqlStr().replace("?",cellValue.toString());
                List<Map<String,Object>> obj= jdbcTemplate.queryForList(key.sqlStr(),cellValue.toString());
                if(obj == null || obj.size() == 0){
                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }
                    cell.setCellComment(setComment(sheet,"数据不存在" ));
                    styleDate.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    cell.setCellStyle(styleDate);
                    throw new RuntimeException("数据不存在");
                }
                cellValue =obj.get(0).get("value");
            }
            break;
        }
        return cellValue;
    }

    /**
     *  比较两个值的大小
     * @param v1 值一
     * @param v2 值二
     * @return 比较的结果
     */
    private boolean compareValue(String v1,String v2) {
        Double d1 = Double.parseDouble(v1);
        Double d2 = Double.parseDouble(v2);
        return d1 >= d2;
    }

    /**
     *
     * @param sheet sheet页
     * @param msg 信息
     */
    abstract Comment setComment(Sheet sheet, String msg);

    /**
     *
     * @param sheet sheet页
     * @param msg 信息
     * @param colIndex 列数
     * @param rowIndex 行数
     */
    abstract Comment setCommentByAnchor(Sheet sheet, String msg, int colIndex, int rowIndex);

    abstract Comment setCommentByAnchorSingle(Sheet sheet, String msg, int colIndex, int rowIndex);

    /**
     判断指定的单元格是否是合并单元格
     * @param sheet sheet页
     * @param row 行下标
     * @param column 列下标
     * @return 返回布尔值
     * */
    private boolean isMergedRegion(Sheet sheet, int row , int column) {
        int sheetMergeCount = sheet.getNumMergedRegions();
        for (int i = 0; i < sheetMergeCount; i++) {
            CellRangeAddress range = sheet.getMergedRegion(i);
            int firstColumn = range.getFirstColumn();
            int lastColumn = range.getLastColumn();
            int firstRow = range.getFirstRow();
            int lastRow = range.getLastRow();
            if(row >= firstRow && row <= lastRow){
                if(column >= firstColumn && column <= lastColumn){
                    return true;
                }
            }
        }
        return false;
    }

    /**
     * 判断double是否是整数
     * @param obj 数据
     * @return 是否为整数
     */
    private static boolean isIntegerForDouble(double obj) {
        double eps = 1e-10;  // 精度范围
        return obj-Math.floor(obj) < eps;
    }

    public List<?> parse(Workbook workbook, Class<?> clazz) {
        return null;
    }
}
