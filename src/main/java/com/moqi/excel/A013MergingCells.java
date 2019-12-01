package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 合并单元格
 *
 * @author moqi
 * On 12/1/19 16:00
 */
@Slf4j
public class A013MergingCells {

    public static void main(String[] args) {

        Workbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("new sheet");
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        cell.setCellValue("This is a test of merging");
        // 四个参数分别是：起始行、终止行、起始列、终止列，全部以左上角单元格 (0,0) 为绝对坐标
        sheet.addMergedRegion(new CellRangeAddress(
                1,
                10,
                1,
                10
        ));

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);

        log.info("程序执行完毕");

    }

}
