package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.DataConsolidateFunction;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFPivotTable;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 创建数据透视表
 * 数据透视表是电子表格文件的强大功能。您可以使用以下代码创建数据透视表。
 *
 * @author moqi
 * On 12/8/19 00:07
 */

public class A043PivotTable {

    public static void main(String[] args) {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet("one");
        // 创建一些数据以在其上构建数据透视表
        for (int i = 0; i < 10; i++) {
            Row row = sheet.createRow(i);
            for (int j = 0; j < 10; j++) {
                // 写入 String 值
                row.createCell(j).setCellValue(String.valueOf(i + j));
            }
            sheet.autoSizeColumn(i);
        }


        XSSFPivotTable pivotTable = sheet.createPivotTable(new AreaReference("A1:D4", null), new CellReference("B2"));
        // 配置数据透视表
        // 使用第一列作为行标签
        pivotTable.addRowLabel(0);
        // 总结第二列
        pivotTable.addColumnLabel(DataConsolidateFunction.SUM, 1);
        // 将第三列设置为过滤器
        pivotTable.addColumnLabel(DataConsolidateFunction.AVERAGE, 2);
        // 在第四列添加过滤器
        pivotTable.addReportFilter(3);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
