package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 使用字体
 * <p>
 * 请注意，工作簿中唯一字体的最大数量限制为32767。
 * 您应该在应用程序中重新使用字体，而不是为每个单元格创建字体。
 *
 * @author moqi
 * On 12/1/19 16:09
 */
@Slf4j
public class A014WorkingWithFonts {

    private static final String TIMES_NEW_ROMAN = "Times New Roman";
    private static final int TEN_THOUSAND = 10000;

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

        Sheet wrongWayToUseFont = wb.createSheet("Wrong Way to use Font");
        wrongWay(wb, wrongWayToUseFont);

        Sheet rightWayToUseFont = wb.createSheet("Right Way to use Font");
        rightWay(wb, rightWayToUseFont);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);

        log.info("程序执行完毕");

    }

    /**
     * 错误的方式使用字体对象，这会创建一万个字体对象
     *
     * @param workbook workbook
     * @param sheet    sheet
     */
    private static void wrongWay(Workbook workbook, Sheet sheet) {
        for (int i = 0; i < TEN_THOUSAND; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            CellStyle style = workbook.createCellStyle();
            Font font = workbook.createFont();
            font.setFontName(TIMES_NEW_ROMAN);
            style.setFont(font);
            cell.setCellStyle(style);
        }
    }

    /**
     * 正确的方式使用字体对象
     *
     * @param workbook workbook
     * @param sheet    sheet
     */
    private static void rightWay(Workbook workbook, Sheet sheet) {
        CellStyle style = workbook.createCellStyle();
        Font font = workbook.createFont();
        font.setFontName(TIMES_NEW_ROMAN);
        style.setFont(font);
        for (int i = 0; i < TEN_THOUSAND; i++) {
            Row row = sheet.createRow(i);
            Cell cell = row.createCell(0);
            cell.setCellStyle(style);
        }
    }

}
