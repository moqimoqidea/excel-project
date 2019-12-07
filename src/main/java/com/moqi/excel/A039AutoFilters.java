package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

/**
 * 自动过滤器
 * 测试创建 Excel 可以，读取 Excel 再操作就失败。
 *
 * @author moqi
 * On 12/7/19 23:23
 */

public class A039AutoFilters {

    private static final String A039_SOURCE_PATH = "A039.xlsx";

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        sheet.setAutoFilter(CellRangeAddress.valueOf("A1:F1"));


        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + A039_SOURCE_PATH);
    }

}
