package com.bzc.example.demo.utils;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import org.springframework.stereotype.Component;

@Component
public class XSSFWorkbookParser extends AbstractWorkBookParser {

    @Override
    CellStyle createCellStyle(Workbook workbook) {
        XSSFWorkbook xssfWorkbook = (XSSFWorkbook) workbook;
        XSSFCellStyle styleDate = xssfWorkbook.createCellStyle();
        styleDate.setFillPattern(FillPatternType.forInt(1));
        styleDate.setBorderTop(BorderStyle.valueOf((short)1));
        styleDate.setBorderLeft(BorderStyle.valueOf((short)1));
        styleDate.setBorderRight(BorderStyle.valueOf((short)1));
        styleDate.setBorderBottom(BorderStyle.valueOf((short)1));
        return styleDate;
    }

    @Override
    Comment setComment(Sheet sheet, String msg) {
        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        XSSFDrawing p = xssfSheet.getDrawingPatriarch();

        if (p == null)
            p = xssfSheet.createDrawingPatriarch();

        XSSFComment comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short)3, 0, (short)5, 6));
        comment.setString(new XSSFRichTextString(msg));
        return comment;
    }

    @Override
    Comment setCommentByAnchor(Sheet sheet, String msg, int colIndex, int rowIndex) {
        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        XSSFDrawing p = xssfSheet.getDrawingPatriarch();

        if (p == null)
            p = xssfSheet.createDrawingPatriarch();

        XSSFComment comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short)(colIndex), rowIndex, (short)(colIndex + 3), rowIndex + 3));

        comment.setString(new XSSFRichTextString(msg));
        return comment;
    }

    Comment setCommentByAnchorSingle(Sheet sheet, String msg, int colIndex, int rowIndex) {
        XSSFSheet xssfSheet = (XSSFSheet) sheet;
        XSSFDrawing p = xssfSheet.getDrawingPatriarch();

        if (p == null)
            p = xssfSheet.createDrawingPatriarch();

        XSSFComment comment = p.createCellComment(new XSSFClientAnchor(0, 0, 0, 0, (short)(colIndex), rowIndex, (short)(colIndex), rowIndex));
        comment.setString(new XSSFRichTextString(msg));

        return comment;
    }
}
