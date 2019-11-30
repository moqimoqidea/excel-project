package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

/**
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

        try {
            Tool.removeOldFile(TARGET_PATH);
            FileOutputStream out = new FileOutputStream(new File(TARGET_PATH));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            log.warn("A002FirstLoadFileExcel main 方法发生异常");
        }

        log.info("程序执行完毕");
    }

}
