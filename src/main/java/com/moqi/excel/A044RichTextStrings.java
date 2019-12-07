package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRichTextString;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.*;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;
import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 具有多种样式的单元格（富文本字符串）
 * 要将一组文本格式（颜色，样式，字体等）应用于单元格，则应 为工作簿创建一个 CellStyle，然后将其应用于单元格。
 * <p>
 * 要将不同的格式应用于单元格的不同部分，您需要使用 RichTextString，它允许样式化单元格中文本的各个部分。
 * https://poi.apache.org/apidocs/dev/org/apache/poi/ss/usermodel/RichTextString.html
 * <p>
 * HSSF和XSSF之间有一些细微的差异，特别是在字体颜色（两种格式在内部存储颜色方面非常不同）方面，
 * 请参见 HSSF Rich Text String 和 XSSF Rich Text String javadocs了解更多详细信息。
 * https://poi.apache.org/apidocs/dev/org/apache/poi/hssf/usermodel/HSSFRichTextString.html
 * https://poi.apache.org/apidocs/dev/org/apache/poi/xssf/usermodel/XSSFRichTextString.html
 *
 * @author moqi
 * On 12/8/19 00:15
 */

public class A044RichTextStrings {

    public static void main(String[] args) {

        hssf();

        xssf();

    }

    private static void hssf() {
        Workbook wb = new HSSFWorkbook();
        Sheet one = wb.createSheet("one");
        Row row = one.createRow(0);
        // HSSF Example
        HSSFCell hssfCell = (HSSFCell) row.createCell(1);

        //rich text consists of two runs
        HSSFRichTextString richString = new HSSFRichTextString("Hello, World!");
        Font font = wb.createFont();
        font.setFontName("Microsoft YaHei");
        richString.applyFont(0, 6, font);

        Font font2 = wb.createFont();
        font2.setFontHeightInPoints((short) 18);
        richString.applyFont(6, 13, font2);
        hssfCell.setCellValue(richString);

        Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);

    }

    private static void xssf() {
        Workbook wb = new XSSFWorkbook();
        Sheet one = wb.createSheet("one");
        Row row = one.createRow(0);

        // XSSF Example
        XSSFCell cell = (XSSFCell) row.createCell(1);
        XSSFRichTextString rt = new XSSFRichTextString("The quick brown fox");
        XSSFFont font1 = (XSSFFont) wb.createFont();
        font1.setBold(true);
        font1.setColor(new XSSFColor(new java.awt.Color(255, 0, 0)));
        rt.applyFont(0, 10, font1);
        XSSFFont font2 = (XSSFFont) wb.createFont();
        font2.setItalic(true);
        font2.setUnderline(XSSFFont.U_DOUBLE);
        font2.setColor(new XSSFColor(new java.awt.Color(0, 255, 0)));
        rt.applyFont(10, 19, font2);
        XSSFFont font3 = (XSSFFont) wb.createFont();
        font3.setColor(new XSSFColor(new java.awt.Color(0, 0, 255)));
        rt.append(" Jumped over the lazy dog", font3);
        cell.setCellValue(rt);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
