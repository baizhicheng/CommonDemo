package com.bzc.example.demo.constants;

import java.math.BigDecimal;
import java.util.HashMap;
import java.util.Map;
import java.util.Random;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class GlobalConstants {

    public static final int INT_0 = 0;

    public static final int INT_1 = 1;

    public static final int INT_2 = 2;

    public static final int INT_3 = 3;

    public static final int INT_100 = 100;

    public static final int INT_200 = 200;

    public static final int INT_500 = 500;

    public static String STRING_0 = "0";

    /**
     * ENCODING_CHARSET
     */
    public static final String ENCODING_CHARSET = "UTF-8";

    /**
     * DATE_FORMAT
     */
    public static final String DATE_FORMAT = "yyyy-MM-dd";

    /**
     * DATE_FORMAT_FULL
     */
    public static final String DATE_FORMAT_FULL = "yyyy-MM-dd HH:mm:ss";

    /**
     * DATE_FORMAT_COUNT
     */
    public static final String DATE_FORMAT_COUNT = "yyyy-MM-01 00:00:00";

    /**
     * TIME_FORMAT
     */
    public static final String TIME_FORMAT = "yyyyMMddHHmmss";

    /**
     * 状态为成功
     */
    public static final String RETURN_SUCCESS = "1";

    /**
     * 状态为失败
     */
    public static final String RETURN_DENIED = "0";

    /**
     * 成功的状态码
     */
    public static final String STATUS_SUCCESS = "200";

    /**
     * 查询成功
     */
    public static final String QUERY_SUCCESS_MSG = "查询成功";

    /**
     * 查询失败
     */
    public static final String QUERY_FAILED_MSG = "查询失败";

    /**
     * 导出成功
     */
    public static final String EXPORT_SUCCESS_MSG = "导出成功";

    /**
     * 导出失败
     */
    public static final String EXPORT_FAILED_MSG = "导出失败";

    /**
     * 更新
     */
    public static final String UPDATE = "update";

    /**
     * 导入
     */
    public static final String IMPORT = "import";

    /**
     * SERVER_ERROR_CODE
     */
    public static final String SERVER_ERROR_CODE = "500";

    /**
     * FIELD_ID
     */
    public static final String FIELD_ID = "id";


    public static final String MSG_EXCEL_ERROR = "导入模板格式错误";
    public static final String MSG_FILE_TYPE_ERROR = "文件类型错误";
    public static final String MSG_EXCEL_DATA_ERROR = "导入数据异常";
    public static final String MSG_EXCEL_DATA_SUCCESS = "导入数据成功";
    public static final String MSG_EXCEL_SUCCESS = "导入成功";
    public static final String MSG_EXCEL_FIAL = "导入失败";
    public static final String MSG_ERROR_DATA = "参数异常";
    public static final String MSG_EXCEL_EXPORT_SUCCESS = "导出成功";
    public static final String MSG_EXCEL_EXPORT_ERROR = "导出失败";
    public static final String MSG_EXCEL_EXIST = "导入数据已存在，是否覆盖？";
    public static final String MSG_EXCEL_PLACEON = "导入数据已归档，是否更新预估值？";
    public static final String MSG_EXCEL_AREAAUTH = "没有相应地区操作权限";

    /**
     * 过滤特殊字符
     */
    private static final String REG_FILTER = "[`~!@#$%^&*()+=|{}':;,\\[\\].<>/?！￥…（）—【】‘；：”“’。，、？]";

    /**
     * 逗号
     */
    public static final String COMMA = ",";

    public static String getRegString(String context) {
        Pattern p = Pattern.compile(REG_FILTER);
        Matcher m = p.matcher(context);
        return m.replaceAll(COMMA).trim();
    }

    /**
     * 每个SHEET页的记录数
     */
    public static final int RECORDS_MAX_NUMBER = 50000;

    /**
     * 时间前区间

     */
    public static final String TIME_FRONT_SECTION = "时间前区间";

    /**
     *   时间后区间
     */
    public static final String TIME_AFTER_SECTION = "时间后区间";

    public static final String IMPORT_CITY_COUNTY_ERROR = "参数错误，请填写数字！";
    public static final String IMPORT_PLAIN_HILL_ERROR = "参数错误，请填写数字！";
    public static final String IMPORT_MOUNTAIN_ISLAND_ERROR = "参数错误，请填写数字！";


    /**
     * 导出Excel标题
     */
    public static final String[] SHEET_EXPORT_TITLE = {"ID","用户名","年龄"};

    /**
     * 导入Excel标题
     */
    public static final String[] SHEET_IMPORT_TITLE = {"ID","用户名","年龄"};

    public static final BigDecimal ZERO = new BigDecimal(0);
    public static final BigDecimal DAYS_COUNT_OF_YEAR = new BigDecimal(365);
    public static final BigDecimal DAYS_COUNT_OF_MONTH = new BigDecimal(30);
    public static final BigDecimal EIGHTY = new BigDecimal(80);
    public static final BigDecimal DEFAULT = new BigDecimal(new Random().nextInt(5)+80);

    /**
     * 时间yyyy-MM-dd HH:mm:ss格式校验
     */
    public static final String DATE_TIME_REP = "((^[1-9]\\d{3}-(0[1-9]|1[0-2])-(0[1-9]|[1-2][0-9]|3[0-1])\\s+(20|21|22|23|[0-1]\\d):[0-5]\\d:[0-5]\\d$))";

    // 最多三位小数
    public static final String THREE_DECIMAL = "^\\d+(\\.\\d{1,3})?$";

}
