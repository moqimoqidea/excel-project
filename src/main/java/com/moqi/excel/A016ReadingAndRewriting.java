package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.*;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 读取和重写 workbook
 *
 * @author moqi
 * On 12/1/19 17:03
 */

public class A016ReadingAndRewriting {

    public static void main(String[] args) throws IOException {

        try (InputStream inp = new FileInputStream(DEFAULT_XLS_PATH)) {
            Workbook wb = WorkbookFactory.create(inp);
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(8);
            Cell cell = row.getCell(3);

            if (cell == null) {
                cell = row.createCell(3);
            }

            cell.setCellValue("a test");

            Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);
        }

    }

}
