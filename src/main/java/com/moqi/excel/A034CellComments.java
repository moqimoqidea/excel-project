package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.Map;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

/**
 * 单元格注释-HSSF和XSSF
 * 注释是附于单元格并与单元格相关联的富文本注释，与其他单元格内容分开。
 * 注释内容与单元格分开存储，并显示在与单元格独立但关联的图形对象（如文本框）中
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A034CellComments {


    private static final String COMMENT_XSSF_XLSX = "comment-xssf.xlsx";

    public static void main(String[] args) throws IOException {
        // 创建单元格注释
        createCellComment();

        // 读取固定位置单元格注释
        readOneCellComment();

        // 读取所有单元格注释
        readAllCellComment();
    }

    private static void createCellComment() {
        Workbook wb = new XSSFWorkbook();
        CreationHelper factory = wb.getCreationHelper();
        Sheet sheet = wb.createSheet("Sheet1");
        Row row = sheet.createRow(3);
        Cell cell = row.createCell(5);
        cell.setCellValue("F4");
        Drawing<?> drawing = sheet.createDrawingPatriarch();

        // 当注释框可见时，将其显示在1x3的空间中
        ClientAnchor anchor = factory.createClientAnchor();
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + 1);
        anchor.setRow1(row.getRowNum());
        anchor.setRow2(row.getRowNum() + 3);

        // 创建评论并设置文字+作者
        Comment comment = drawing.createCellComment(anchor);
        RichTextString str = factory.createRichTextString("Hello, World!");
        comment.setString(str);
        comment.setAuthor("Apache POI");

        // 将注释分配给单元格
        cell.setCellComment(comment);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + COMMENT_XSSF_XLSX);
    }

    private static void readOneCellComment() throws IOException {
        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLSX_DIR_PATH + COMMENT_XSSF_XLSX));
        Sheet sheet = wb.getSheet("Sheet1");

        Cell cell = sheet.getRow(3).getCell(5);
        Comment comment = cell.getCellComment();
        if (comment != null) {
            RichTextString str = comment.getString();
            String author = comment.getAuthor();

            log.info("注释的内容是:{},作者是:{}", str, author);
        }
        //  或者，您也可以按（行，列）检索单元格注释
        Comment comment2 = sheet.getCellComment(new CellAddress(3, 5));
        if (comment2 != null) {
            log.info("注释的内容是:{},作者是:{},被注释的主体是:{}", comment2.getString(), comment2.getAuthor(), comment2.getAddress());
        }
    }

    private static void readAllCellComment() throws IOException {
        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLSX_DIR_PATH + COMMENT_XSSF_XLSX));
        Sheet sheet = wb.getSheet("Sheet1");

        Map<CellAddress, ? extends Comment> comments = sheet.getCellComments();
        for (Map.Entry<CellAddress, ? extends Comment> e : comments.entrySet()) {
            CellAddress loc = e.getKey();
            Comment comment = e.getValue();
            log.info("Comment at {}: [{}] {}", loc, comment.getAuthor(), comment.getString().getString());
        }


    }


}
