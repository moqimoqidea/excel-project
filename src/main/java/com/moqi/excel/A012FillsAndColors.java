package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 填充单元格的颜色：分为前景色和背景色
 *
 * @author moqi
 * On 12/1/19 15:48
 */
@Slf4j
public class A012FillsAndColors {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        // 创建一行并在其中放入一些单元格。行从0开始。
        Row row = sheet.createRow(1);

        // 红色背景色
        CellStyle style = wb.createCellStyle();
        style.setFillBackgroundColor(IndexedColors.RED.getIndex());
        style.setFillPattern(FillPatternType.BIG_SPOTS);
        Cell cell = row.createCell(1);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        // 橙色 单元格前景色，前景色不是字体颜色。
        style = wb.createCellStyle();
        style.setFillForegroundColor(IndexedColors.ORANGE.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell = row.createCell(2);
        cell.setCellValue("X");
        cell.setCellStyle(style);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);

        log.info("程序执行完毕");
    }

}
