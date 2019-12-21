package com.bzc.example.demo.utils;


import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.*;
import org.springframework.stereotype.Component;

@Component
public class HSSFWorkbookParser extends AbstractWorkBookParser{

    @Override
    CellStyle createCellStyle(Workbook workbook) {
        HSSFWorkbook hssfWorkbook = (HSSFWorkbook) workbook;
        HSSFCellStyle styleDate = hssfWorkbook.createCellStyle();
        styleDate.setFillPattern(FillPatternType.forInt(1));
        styleDate.setBorderTop(BorderStyle.NONE);
        styleDate.setBorderTop(BorderStyle.valueOf((short)1));
        styleDate.setBorderLeft(BorderStyle.valueOf((short)1));
        styleDate.setBorderRight(BorderStyle.valueOf((short)1));
        styleDate.setBorderBottom(BorderStyle.valueOf((short)1));
        return styleDate;
    }

    @Override
    Comment setComment(Sheet sheet, String msg) {
        HSSFSheet hssfSheet = (HSSFSheet) sheet;
        HSSFPatriarch p = hssfSheet.createDrawingPatriarch();
        HSSFComment comment = p.createComment(new HSSFClientAnchor(0, 0, 0, 0, (short)3, 0, (short)5, 6));
        comment.setString(new HSSFRichTextString(msg));
        return comment;
    }

    @Override
    Comment setCommentByAnchor(Sheet sheet, String msg, int colIndex, int rowIndex) {
        HSSFSheet hssfSheet = (HSSFSheet) sheet;
        HSSFPatriarch p = hssfSheet.getDrawingPatriarch();

        if (p == null)
            p = hssfSheet.createDrawingPatriarch();

        HSSFComment comment = p.createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short)(colIndex + 1), rowIndex - 1, (short)(colIndex + 3), rowIndex + 5));
        comment.setString(new HSSFRichTextString(msg));
        return comment;
    }

    @Override
    Comment setCommentByAnchorSingle(Sheet sheet, String msg, int colIndex, int rowIndex) {
        HSSFSheet hssfSheet = (HSSFSheet) sheet;
        HSSFPatriarch p = hssfSheet.getDrawingPatriarch();

        if (p == null)
            p = hssfSheet.createDrawingPatriarch();

        HSSFComment comment = p.createCellComment(new HSSFClientAnchor(0, 0, 0, 0, (short)(colIndex), rowIndex, (short)(colIndex), rowIndex));
        comment.setString(new HSSFRichTextString(msg));
        return comment;
    }
}
