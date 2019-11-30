package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_EXCEL_PATH;

/**
 * 初始化空的 Excel
 *
 * @author moqi
 * On 11/30/19 21:33
 */
@Slf4j
public class A001BlankExcel {


    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Demo Sheet");
        sheet.createRow(0).createCell(0).setCellValue("Demo Cell");

        Tool.generateExcelFile(workbook, DEFAULT_EXCEL_PATH);

        log.info("程序执行完毕");
    }


}
