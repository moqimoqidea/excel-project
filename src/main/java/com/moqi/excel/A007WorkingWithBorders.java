package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_EXCEL_PATH;

/**
 * 设置单元框的格式
 *
 * @author moqi
 * On 12/1/19 10:56
 */

public class A007WorkingWithBorders {

    public static void main(String[] args) {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        // Create a row and put some cells in it. Rows are 0 based.
        Row row = sheet.createRow(1);
        // Create a cell and put a value in it.
        Cell cell = row.createCell(1);
        cell.setCellValue(4);
        // 用周围的边框设置单元格的样式。
        CellStyle style = wb.createCellStyle();
        style.setBorderBottom(BorderStyle.THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(BorderStyle.THIN);
        style.setLeftBorderColor(IndexedColors.GREEN.getIndex());
        style.setBorderRight(BorderStyle.THIN);
        style.setRightBorderColor(IndexedColors.BLUE.getIndex());
        style.setBorderTop(BorderStyle.MEDIUM_DASHED);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        cell.setCellStyle(style);

        Tool.generateExcelFile(wb, DEFAULT_EXCEL_PATH);

    }


}
