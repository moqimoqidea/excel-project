package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileOutputStream;

/**
 * 初始化空的 Excel
 *
 * @author moqi
 * On 11/30/19 21:33
 */
@Slf4j
public class A001BlankExcel {

    private static final String PATH = "/Users/moqi/Downloads/excel_test/demo.xlsx";

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Demo Sheet");
        sheet.createRow(0).createCell(0).setCellValue("Demo Cell");

        try {
            Tool.removeOldFile(PATH);
            FileOutputStream out = new FileOutputStream(new File(PATH));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            log.warn("A001BlankExcel main 方法发生异常");
        }

        log.info("程序执行完毕");
    }


}
