package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 在工作表上上移或下移行
 * 将已完成的工作表个别行进行移动
 *
 * @author moqi
 * On 12/1/19 19:42
 */
@Slf4j
public class A023ShiftRowsUpOrDownOnSheet {

    private static final int FIFTEEN = 15;

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet = wb.createSheet("row sheet");
        // 为电子表格创建各种单元格和行。
        for (int i = 0; i < FIFTEEN; i++) {
            // 使用 CellUtil 创建值
            Cell cell = CellUtil.createCell(sheet.createRow(i), 0, "Row is " + i);
            // 使用 CellUtil 对 cell 进行居中
            CellUtil.setAlignment(cell, HorizontalAlignment.CENTER);
        }
        // 将电子表格上的6-11行移至顶部（第0-5行）
        sheet.shiftRows(5, 10, -5);
        // // 自动调整列宽以适合内容
        sheet.autoSizeColumn(0);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
