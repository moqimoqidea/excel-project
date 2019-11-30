package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Calendar;
import java.util.Date;

import static com.moqi.constant.Constant.DEFAULT_EXCEL_PATH;

/**
 * 处理不同类型的单元格
 *
 * @author moqi
 * On 11/30/19 22:15
 */
@Slf4j
public class A003DifferentType {

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("A003DifferentType Sheet");
        XSSFRow row = sheet.createRow(2);
        row.createCell(0).setCellValue(1);
        row.createCell(1).setCellValue(2.2);
        row.createCell(2).setCellValue(new Date());
        row.createCell(3).setCellValue(Calendar.getInstance());
        row.createCell(4).setCellValue("A String");
        row.createCell(5).setCellValue(true);
        row.createCell(6).setCellType(CellType.ERROR);

        Tool.getExcelFile(workbook, DEFAULT_EXCEL_PATH);

        log.info("程序执行完毕");
    }

}
