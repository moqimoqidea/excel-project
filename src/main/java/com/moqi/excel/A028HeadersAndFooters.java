package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFHeader;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;
import static com.moqi.tool.Tool.fill100ValueOnSheet;

/**
 * 页眉和页脚
 * 示例适用于页眉，但直接应用于页脚。
 * <p>
 * 本例使用 xls
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A028HeadersAndFooters {


    public static void main(String[] args) {
        Workbook wb = new HSSFWorkbook();

        Sheet sheet1 = wb.createSheet("Sheet1");

        fill100ValueOnSheet(sheet1);

        Header header = sheet1.getHeader();
        header.setCenter("Center Header");
        header.setLeft("Left Header");
        header.setRight(HSSFHeader.font("Stencil-Normal", "Italic") +
                HSSFHeader.fontSize((short) 16) + "Right w/ Stencil-Normal Italic font and size 16");

        Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);
    }


}
