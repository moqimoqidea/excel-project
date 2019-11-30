package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_PATH;

/**
 * 展示各种对齐方式
 *
 * @author moqi
 * On 11/30/19 23:07
 */

public class A006DemonstratesVariousAlignmentOptions {

    private static final String XSSF_ALIGN_XLSX = "xssf-align.xlsx";

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet();
        Row row = sheet.createRow(2);
        row.setHeightInPoints(30);

        createCell(wb, row, 0, HorizontalAlignment.CENTER, VerticalAlignment.BOTTOM);
        createCell(wb, row, 1, HorizontalAlignment.CENTER_SELECTION, VerticalAlignment.BOTTOM);
        createCell(wb, row, 2, HorizontalAlignment.FILL, VerticalAlignment.CENTER);
        createCell(wb, row, 3, HorizontalAlignment.GENERAL, VerticalAlignment.CENTER);
        createCell(wb, row, 4, HorizontalAlignment.JUSTIFY, VerticalAlignment.JUSTIFY);
        createCell(wb, row, 5, HorizontalAlignment.LEFT, VerticalAlignment.TOP);
        createCell(wb, row, 6, HorizontalAlignment.RIGHT, VerticalAlignment.TOP);
        createCell(wb, row, 7, HorizontalAlignment.CENTER, VerticalAlignment.CENTER);

        Tool.generateExcelFile(wb, DEFAULT_PATH + XSSF_ALIGN_XLSX);
    }

    /**
     * 创建一个单元格并以某种方式对齐它。
     *
     * @param wb                  the workbook
     * @param row                 the row to create the cell in
     * @param column              the column number to create the cell in
     * @param horizontalAlignment 将单元格的水平对齐方式对齐
     * @param verticalAlignment   将单元格的垂直对齐方式
     */
    private static void createCell(Workbook wb,
                                   Row row,
                                   int column,
                                   HorizontalAlignment horizontalAlignment,
                                   VerticalAlignment verticalAlignment) {
        Cell cell = row.createCell(column);
        cell.setCellValue("Align It");
        CellStyle cellStyle = wb.createCellStyle();
        cellStyle.setAlignment(horizontalAlignment);
        cellStyle.setVerticalAlignment(verticalAlignment);
        cell.setCellStyle(cellStyle);
    }

}
