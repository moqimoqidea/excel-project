package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;
import static com.moqi.tool.Tool.fillValueOnSheet;

/**
 * 打印时重复特定行和列
 * <p>
 * 通过使用Sheet类中的 setRepeatingRows() 和 setRepeatingColumns() 方法，
 * 可以在打印输出中设置重复的行和列。
 * <p>
 * 这些方法需要一个CellRangeAddress参数，该参数指定要重复的行或列的范围。
 * 对于setRepeatingRows()，它应指定要重复的行范围，其中列部分跨越所有列。
 * 对于setRepeatingColumns()，应指定要重复的列范围，其中行部分跨越所有行。
 * 如果参数为null，则将删除重复的行或列。
 * <p>
 * 打印时起作用：调节的属性是
 * File -> Page Setup -> Sheet -> Print Titles 中的
 * Rows to repeat at top:
 * 和
 * Columns to repeat at left:
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A027RepeatingRowsAndColumns {

    private static final String A027_PATH = "A027.xlsx";


    public static void main(String[] args) throws IOException, InvalidFormatException {
        Workbook wb = new XSSFWorkbook(new File(DEFAULT_XLSX_DIR_PATH + A027_PATH));

        Sheet sheet1 = wb.getSheet("Sheet1");
        Sheet sheet2 = wb.getSheet("Sheet2");

        fillValueOnSheet(sheet1);
        fillValueOnSheet(sheet2);

        // 将行设置为在第一张纸上的第4到5行重复。
        sheet1.setRepeatingRows(CellRangeAddress.valueOf("4:5"));
        // 将列设置为在第二张纸上从A列重复到C列
        sheet2.setRepeatingColumns(CellRangeAddress.valueOf("A:C"));

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + A027_PATH);
    }


}
