package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Calendar;
import java.util.Date;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 将日期正确放在单元格中
 *
 * @author moqi
 * On 11/30/19 22:15
 */
@Slf4j
public class A004DateFormat {

    public static void main(String[] args) {

        XSSFWorkbook workbook = new XSSFWorkbook();
        CellStyle ymdHmsCellStyle = Tool.getYmdHmsCellStyle(workbook);
        CellStyle ymdCellStyle = Tool.getYmdCellStyle(workbook);

        XSSFSheet sheet = workbook.createSheet("A004DateFormat Sheet");
        XSSFRow row = sheet.createRow(0);

        XSSFCell defaultCell = row.createCell(0);
        defaultCell.setCellValue(new Date());

        XSSFCell newDateCell = row.createCell(1);
        newDateCell.setCellValue(new Date());
        newDateCell.setCellStyle(ymdHmsCellStyle);

        XSSFCell calenderCell = row.createCell(2);
        calenderCell.setCellValue(Calendar.getInstance());
        calenderCell.setCellStyle(ymdHmsCellStyle);

        XSSFCell calenderYmdCell = row.createCell(3);
        calenderYmdCell.setCellValue(Calendar.getInstance());
        calenderYmdCell.setCellStyle(ymdCellStyle);

        Tool.generateExcelFile(workbook, DEFAULT_XLSX_PATH);

        log.info("程序执行完毕");
    }

}
