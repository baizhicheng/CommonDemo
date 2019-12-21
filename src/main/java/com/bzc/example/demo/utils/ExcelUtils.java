package com.bzc.example.demo.utils;

import com.bzc.example.demo.vo.CommentVo;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.stereotype.Component;

import javax.servlet.http.HttpServletRequest;
import java.io.*;
import java.net.URLEncoder;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.Random;

@Slf4j
@Component
public class ExcelUtils {

    public static final String EXCEL_FIELD_INDEX = "excelFieldIndex";
    // Excel数据总数
    public static String TOTAL_COUNT = "totalCount";
    // Excel非法数据总数
    public static String INVALID_COUNT = "invalidCount";
    // 合法数据list
    public static String VALID_LIST = "validList";
    // 操作码，按毫秒记录
    public static String OPERATION_CODE = "operationCode";
    // workbook
    public static String WORK_BOOK = "workBook";

    public static String FILEPATH_CACHENAME  = "filePath:";
    // 新增数量
    public static String INSERT_COUNT = "insertCount";
    // 更新数量
    public static String UPDATE_COUNT = "updateCount";
    // 文件名
    public static String FILE_NAME = "fileName";
    // 单元格索引
    public static String ROW_INDEX = "rowIndex";

    public static String SAVE_FILE = "saveFile";

    public static String SHEET_INDEX = "sheetIndex";
    // 导入文件的标题
    public static String SHEET_TITLES = "sheetTitles";
    // 模板标题错误
    public static String SHEET_TITLE_ERROR = "sheetTitleError";
    // 是否导入成功
    public static String IS_SUCCESS = "isSuccess";

    public static String MSG = "msg";

    @Autowired
    private HSSFWorkbookParser hssfWorkbookParser;

    @Autowired
    private XSSFWorkbookParser xssfWorkbookParser;

    public List<?> parse(InputStream is, Class<?> clazz) throws Exception {
        Workbook workbook = WorkbookFactory.create(is);
        if (workbook instanceof HSSFWorkbook) {
            return  hssfWorkbookParser.parse(workbook,clazz);
        } else if (workbook instanceof XSSFWorkbook) {
            return xssfWorkbookParser.parse(workbook,clazz);
        }
        return null;
    }

    public Map<String,Object> parseExcel(Workbook workbook, String fileName, Class<?> clazz) throws Exception {
        Map<String,Object> excel = null;

        if (workbook instanceof HSSFWorkbook) {
            excel = hssfWorkbookParser.parseExcel(workbook,fileName,clazz);
        } else if (workbook instanceof XSSFWorkbook) {
            excel = xssfWorkbookParser.parseExcel(workbook,fileName,clazz);
        }

        return excel;
    }

    public Map<String,Object> parseExcel(Workbook workbook, int sheetIndex, String fileName, Class<?> clazz) throws Exception {
        Map<String,Object> excel = null;

        if (workbook instanceof HSSFWorkbook) {
            excel = hssfWorkbookParser.parseExcel(workbook, sheetIndex, fileName,clazz);
        } else if (workbook instanceof XSSFWorkbook) {
            excel = xssfWorkbookParser.parseExcel(workbook, sheetIndex, fileName,clazz);
        }

        return excel;
    }


    /**
     *  更新导入结果
     * @param map
     * @param insertCount 新增记录数
     * @param updateCount 更新记录数
     * @param errorMap 错误信息
     * @param request
     */
    public Map<String,Object> completeImport(Map<String, Object> map, int insertCount, int updateCount,
                                             Map<Integer, Object> errorMap, HttpServletRequest request) {
        Workbook workbook = (Workbook) map.get(SAVE_FILE);
        int errorCount = 0;

        if (errorMap == null || errorMap.size() == 0) {

        } else {
            errorCount = errorMap.size();
            if (workbook instanceof HSSFWorkbook) {
                hssfWorkbookParser.addErrorComment(workbook,map,errorMap);
            } else if (workbook instanceof XSSFWorkbook) {
                xssfWorkbookParser.addErrorComment(workbook,map,errorMap);
            }
        }

        int  invalidTotal = (int) map.get(INVALID_COUNT) + errorCount;

        map.put(INVALID_COUNT,invalidTotal);
        map.put(ExcelUtils.INSERT_COUNT,insertCount);
        map.put(ExcelUtils.UPDATE_COUNT,updateCount);
        map.remove(SAVE_FILE);
        map.remove(ExcelUtils.VALID_LIST);

        if (invalidTotal > 0) {
            // 保存文件
            String wbKey = ExcelUtils.WORK_BOOK  + "_" + String.valueOf(System.currentTimeMillis());
            map.put(ExcelUtils.WORK_BOOK,wbKey);
            String fileName = (String) map.get(FILE_NAME);
            String savePath = request.getSession().getServletContext().getRealPath("") + "\\excel\\" + wbKey + "\\";
            File pathFile = new File(savePath);

            if (!pathFile.exists())
                pathFile.mkdirs();

//            File savedFile = new File(pathFile, encodeNameByBrowser(fileName,request));

            Date time = new Date();
            SimpleDateFormat formatter = new SimpleDateFormat("yyyyMMddHHmmss");

            Random random = new Random();

            int randNum = random.nextInt(100);
            randNum = 100 + randNum;

            File savedFile = new File(pathFile, formatter.format(time) + randNum + ".xlsx");
            OutputStream outputStream = null;

            try {
                outputStream = new FileOutputStream(savedFile);
                workbook.write(outputStream);
            } catch (Exception e) {
                log.error(e.getMessage(), e);
            } finally {
                try {
                    if (null != outputStream) {
                        outputStream.flush();
                        outputStream.close();
                    }
                } catch (IOException e) {
                    log.error(e.getMessage(), e);
                }
            }
        }
        return map;
    }

    /**
     * 更新导入结果
     *
     * @param is
     * 输入流
     * @param os
     * 输出流
     * @param list
     * 批注信息
     */
    public static void completeImport(InputStream is, OutputStream os, List<CommentVo> list) {
        XSSFWorkbook workBook = null;
        try {
            workBook = new XSSFWorkbook(is);
            XSSFSheet sheet = workBook.getSheetAt(0);
            int rowNum = sheet.getLastRowNum();// 行
            int cellNum = sheet.getRow(0).getLastCellNum();// 列

            for (int i=0;i<rowNum;i++) {
                for (int j=0;j<cellNum;j++) {
                    XSSFRow row = sheet.getRow(i);

                    if(null != row) {
                        XSSFCell cell = row.getCell(j);

                        if(null != cell) {
                            if (null != cell.getCellComment()) {
                                cell.removeCellComment();//删除批注
                            }
                        } else {
                            break;
                        }
                    } else {
                        break;
                    }
                }
            }

            for (CommentVo c : list) {
                // 创建绘图对象
                XSSFDrawing p = sheet.createDrawingPatriarch();
                XSSFRow row = sheet.getRow(c.getRow());

                if (null == row){
                    row = sheet.createRow(c.getRow());
                }

                // 创建单元格对象,批注插入到4行,1列,B5单元格
                XSSFCell cell = row.getCell(c.getCell());

                if (null == cell){
                    cell = row.createCell(c.getCell());
                }

                // 前四个参数是坐标点,后四个参数是编辑和显示批注时的大小.
                XSSFComment comment;

                try {
//                    comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short) 3, 0, (short) 5, 6));
                    comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short)(cellNum), rowNum, (short)(cellNum + 3), rowNum + 3));
                    // 输入批注信息
                    comment.setString(new XSSFRichTextString(c.getMsg()));

                    if (null != cell.getCellComment()) {
                        cell.removeCellComment();// 删除批注
                    }

                    cell.setCellComment(comment);// 把批注赋值给单元格

                    CellStyle tCellStyle = workBook.createCellStyle();
                    tCellStyle.setFillForegroundColor(new HSSFColor.YELLOW().getIndex());
                    tCellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                    cell.setCellStyle(tCellStyle);
                } catch (Exception e) {
                    e.printStackTrace();
                    continue;
                }
            }
            try {
                workBook.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != is) {
                    is.close();
                }
                if (null != os) {
                    os.close();// 关闭文件流
                }
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 创建批注对象
     *
     * @param cell
     * @param row
     * @param msg
     * @return
     */
    public static CommentVo createComment(int row, int cell, String msg) {
        CommentVo comment = new CommentVo();
        comment.setRow(row);// 行
        comment.setCell(cell);// 列
        comment.setMsg(msg);
        return comment;
    }

    /**
     * 下载
     * @param file
     * @param os
     */
    public static void download(File file, OutputStream os) {
        XSSFWorkbook workBook = null;
        InputStream is = null;
        try {
            is = new FileInputStream(file);
            workBook = new XSSFWorkbook(is);
            try {
                workBook.write(os);
            } catch (IOException e) {
                e.printStackTrace();
            }
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            try {
                if (null != os) {
                    os.close();// 关闭文件流
                }
                if (null != is) {
                    is.close();
                }
                file.delete();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }


    /**
     * 根据浏览器编码
     * @param filename
     * @param request
     * @return
     */
    private   String encodeNameByBrowser(String filename, HttpServletRequest request) {
        if (filename == null || "".equals(filename)) return "";
        String rs = "";
        String userAgent = request.getHeader("USER-AGENT");
        try {
            if (userAgent.contains("MSIE")) {// IE浏览器
                rs = URLEncoder.encode(filename, "UTF8");
            } else if (userAgent.contains("Mozilla")) {// google,火狐浏览器
                rs = new String(filename.getBytes(), "ISO8859-1");
            } else {
                rs = URLEncoder.encode(filename, "UTF8");// 其他浏览器
            }
        } catch (Exception e) {
            log.error(e.getMessage(), e);
        }
        return rs;
    }
}