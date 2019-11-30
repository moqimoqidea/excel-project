package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.IOException;

/**
 * 测试使用 Workbook 打开 Excel 文件并写入保存
 * 在来源和目标路径不一致的情况下，来源文件不会变动
 *
 * @author moqi
 * On 11/30/19 21:36
 */
@Slf4j
public class A002FirstLoadFileExcel {

    private static final String PACKAGE_SOURCE_PATH = "src/main/resources/Book1.xlsx";
    private static final String TARGET_PATH = "/Users/moqi/Downloads/excel_test/Book2.xlsx";

    public static void main(String[] args) throws IOException {

        Workbook workbook = WorkbookFactory.create(new File(PACKAGE_SOURCE_PATH));
        Sheet sheet1 = workbook.getSheet("Sheet1");
        Row row = sheet1.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("Just Test");

        Tool.generateExcelFile(workbook, TARGET_PATH);

        log.info("程序执行完毕");
    }

}
