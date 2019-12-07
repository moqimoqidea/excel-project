package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 调整列宽以适合内容
 * sheet.autoSizeColumn（0）; //调整第一列的宽度
 * <p>
 * 仅对于SXSSFWorkbooks，因为随机访问窗口可能会排除工作表中的大多数行，而这是计算列的最佳匹配宽度所必需的，
 * 因此在刷新任何行之前，必须跟踪这些列以自动调整大小。
 * <p>
 * 请注意，Sheet＃autoSizeColumn（）不评估公式单元格，公式单元格的宽度是根据缓存的公式结果计算的。
 * 如果您的工作簿中有很多公式，那么在自动调整大小之前最好对它们进行评估。
 * <p>
 * 警告
 * 要计算列宽，Sheet.autoSizeColumn使用Java2D类，如果图形环境不可用，则该类将引发异常。
 * 如果没有图形环境，则必须告诉Java您正在以无头模式运行，并设置以下系统属性：java.awt.headless = true。
 * 您还应确保在工作簿中使用的字体可用于Java。
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A035ColumnSize {


    private static final int TEN = 10;

    public static void main(String[] args) {

        SXSSFWorkbook workbook = new SXSSFWorkbook();
        SXSSFSheet sheet = workbook.createSheet();
        sheet.trackColumnForAutoSizing(0);
        sheet.trackColumnForAutoSizing(1);
        // 如果您具有列索引的Collection，请参见SXSSFSheet＃trackColumnForAutoSizing（Collection <Integer>）
        // 或滚动自己的for循环。
        // 或者，如果没有自动调整大小的列，请使用SXSSFSheet＃trackAllColumnsForAutoSizing（）
        // 预先知道，或者您正在升级现有代码，并试图最小化更改。记住
        // 由于计算出最适合的宽度，因此跟踪所有列将需要更多的内存和CPU周期
        // 在刷新的每一行的所有跟踪列上。

        // create some cells
        for (int i = 0; i < TEN; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < TEN; j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue("Cell " + cell.getAddress().formatAsString());
            }
        }

        // 自动调整列的大小。
        sheet.autoSizeColumn(0);
        sheet.autoSizeColumn(1);

        Tool.generateExcelFile(workbook, DEFAULT_XLSX_PATH);
    }

}
