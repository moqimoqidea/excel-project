package com.moqi.excel;


import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;


/**
 * Double 类型的格式化，可以精确定位需要几位小数
 *
 * @author moqi
 * On 12/1/19 17:24
 */

public class A018DataFormats {

    private static final String A018_XLSX = "A018.xlsx";

    public static void main(String[] args) {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        CellStyle style;
        DataFormat format = wb.createDataFormat();
        Row row;
        Cell cell;
        int rowNum = 0;
        int colNum = 0;
        row = sheet.createRow(rowNum);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();

        // 保留一位小数
        style.setDataFormat(format.getFormat("0.0"));
        cell.setCellStyle(style);
        rowNum++;
        row = sheet.createRow(rowNum);
        cell = row.createCell(colNum);
        cell.setCellValue(11111.25);
        style = wb.createCellStyle();

        // 保留四位小数
        style.setDataFormat(format.getFormat("#,##0.0000"));
        cell.setCellStyle(style);

        // 自动调整列宽以适合内容
        sheet.autoSizeColumn(0);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + A018_XLSX);
    }

}
