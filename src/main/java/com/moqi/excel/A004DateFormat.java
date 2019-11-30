package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.xssf.usermodel.*;

import java.util.Calendar;
import java.util.Date;

import static com.moqi.constant.Constant.*;

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
        XSSFSheet sheet = workbook.createSheet("A004DateFormat Sheet");
        XSSFRow row = sheet.createRow(0);

        XSSFCell defaultCell = row.createCell(0);
        defaultCell.setCellValue(new Date());

        XSSFCellStyle ymdHmsCellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        ymdHmsCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(YYYY_MM_DD_HH_MM_SS));
        XSSFCell newDateCell = row.createCell(1);
        newDateCell.setCellValue(new Date());
        newDateCell.setCellStyle(ymdHmsCellStyle);

        XSSFCell calenderCell = row.createCell(2);
        calenderCell.setCellValue(Calendar.getInstance());
        calenderCell.setCellStyle(ymdHmsCellStyle);

        XSSFCellStyle ymdCellStyle = workbook.createCellStyle();
        ymdCellStyle.setDataFormat(createHelper.createDataFormat().getFormat(YYYY_MM_DD));
        XSSFCell calenderYmdCell = row.createCell(3);
        calenderYmdCell.setCellValue(Calendar.getInstance());
        calenderYmdCell.setCellStyle(ymdCellStyle);

        Tool.getExcelFile(workbook, DEFAULT_EXCEL_PATH);

        log.info("程序执行完毕");
    }

}
