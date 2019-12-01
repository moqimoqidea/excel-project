package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.ss.util.RegionUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 使用便利功能
 * 便利功能提供实用程序功能，例如在合并区域周围设置边框以及更改样式属性而无需显式创建新样式。
 * 主要是调用 RegionUtil、CellUtil 类的方法
 *
 * @author moqi
 * On 12/1/19 19:30
 */
@Slf4j
public class A022UsingTheConvenienceFunctions {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet1 = wb.createSheet("new sheet");
        // 创建合并区域
        Row row = sheet1.createRow(1);
        Row row2 = sheet1.createRow(2);
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of merging");
        CellRangeAddress region = CellRangeAddress.valueOf("B2:E5");
        sheet1.addMergedRegion(region);

        // 设置边框和边框颜色。
        RegionUtil.setBorderBottom(BorderStyle.MEDIUM_DASHED, region, sheet1);
        RegionUtil.setBorderTop(BorderStyle.MEDIUM_DASHED, region, sheet1);
        RegionUtil.setBorderLeft(BorderStyle.MEDIUM_DASHED, region, sheet1);
        RegionUtil.setBorderRight(BorderStyle.MEDIUM_DASHED, region, sheet1);
        RegionUtil.setBottomBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setTopBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setLeftBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);
        RegionUtil.setRightBorderColor(IndexedColors.AQUA.getIndex(), region, sheet1);

        // 显示 CellUtil 的一些用法
        CellStyle style = wb.createCellStyle();
        style.setIndention((short) 4);
        CellUtil.createCell(row, 8, "This is the value of the cell", style);
        Cell cell2 = CellUtil.createCell(row2, 8, "This is the value of the cell");
        CellUtil.setAlignment(cell2, HorizontalAlignment.CENTER);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
