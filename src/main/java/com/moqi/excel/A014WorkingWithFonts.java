package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 使用字体
 *
 * @author moqi
 * On 12/1/19 16:09
 */
@Slf4j
public class A014WorkingWithFonts {

    private static final String TIMES_NEW_ROMAN = "Times New Roman";

    public static void main(String[] args) {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow(1);

        // 创建一个新字体并进行更改。
        Font font = wb.createFont();
        // 字体大小
        font.setFontHeightInPoints((short) 24);
        // 字体名称
        font.setFontName(TIMES_NEW_ROMAN);
        // 是否斜体
        font.setItalic(true);
        // 字体设置为一种样式，因此请创建一种新样式来使用。
        CellStyle style = wb.createCellStyle();
        style.setFont(font);
        // 创建一个单元格并在其中添加一个值。
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of fonts");
        cell.setCellStyle(style);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);

        log.info("程序执行完毕");

    }

}
