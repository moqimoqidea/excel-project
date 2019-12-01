package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 设置打印区域
 *
 * @author moqi
 * On 12/1/19 19:12
 */
@Slf4j
public class A020SetPrintArea {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        // 设置第一张纸的打印区域
        wb.setPrintArea(0, "$A$1:$C$2");
        // 或者：

        wb.setPrintArea(
                // sheet index
                0,
                // start column
                0,
                // end column
                1,
                // start row
                0,
                // end row
                0
        );

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
