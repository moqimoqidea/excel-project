package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

/**
 * 在单元格中使用换行符
 *
 * @author moqi
 * On 12/1/19 17:10
 */
@Slf4j
public class A017UsingNewlinesInCells {

    private static final String OOXML_NEWLINES_XLSX = "ooxml-newlines.xlsx";

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();

        doubleWidthAndHeight(sheet);

        Row row = sheet.createRow(2);
        Cell cell = row.createCell(2);
        cell.setCellValue("Use \n with word wrap on to create a new line");
        // 要启用换行符，您需要使用wrap = true设置单元格样式
        CellStyle cs = wb.createCellStyle();
        cs.setWrapText(true);
        cell.setCellStyle(cs);
        // 增加行高以容纳两行文本
        row.setHeightInPoints((2 * sheet.getDefaultRowHeightInPoints()));
        // 自动调整列宽以适合内容
        sheet.autoSizeColumn(2);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + OOXML_NEWLINES_XLSX);
    }

    /**
     * 测试 设定 sheet 列宽和行高都为 2 倍
     *
     * @param sheet 表
     */
    private static void doubleWidthAndHeight(Sheet sheet) {
        float defaultRowHeightInPoints = sheet.getDefaultRowHeightInPoints();
        log.info("defaultRowHeightInPoints:{}", defaultRowHeightInPoints);

        int defaultColumnWidth = sheet.getDefaultColumnWidth();
        log.info("defaultColumnWidth:{}", defaultColumnWidth);

        // 设定 sheet 列宽和行高都为 2 倍
        sheet.setDefaultColumnWidth(2 * sheet.getDefaultColumnWidth());
        sheet.setDefaultRowHeightInPoints(2 * sheet.getDefaultRowHeightInPoints());
    }

}
