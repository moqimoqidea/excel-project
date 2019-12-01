package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HeaderFooter;
import org.apache.poi.ss.usermodel.Footer;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 在页脚上设置页码
 *
 * @author moqi
 * On 12/1/19 19:15
 */
@Slf4j
public class A021SetPageNumbersOnFooter {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("format sheet");
        Footer footer = sheet.getFooter();
        footer.setRight("Page " + HeaderFooter.page() + " of " + HeaderFooter.numPages());

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
