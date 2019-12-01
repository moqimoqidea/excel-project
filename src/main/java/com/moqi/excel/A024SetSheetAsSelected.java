package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 将工作表设置为选中状态
 * 201921012148 测试成功了，在 Excel 中选中和激活不是一个概念
 * 正如方法的注释所说：可以选择多个表，但一次只能激活一个表。
 * 打开 Excel 时 firstSheet 默认激活且默认被选择，
 * secondSheet 设定选中所以状态和 thirdSheet 不同。
 * <p>
 * 进阶：切换激活的 sheet 为 fourthSheet
 *
 * @author moqi
 * On 12/1/19 19:49
 */
@Slf4j
public class A024SetSheetAsSelected {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet firstSheet = wb.createSheet("First Sheet");
        Sheet secondSheet = wb.createSheet("Second Sheet");
        Sheet thirdSheet = wb.createSheet("Third Sheet");
        Sheet fourthSheet = wb.createSheet("Fourth Sheet");

        Tool.fill100ValueOnSheet(firstSheet);
        Tool.fill100ValueOnSheet(secondSheet);
        Tool.fill100ValueOnSheet(thirdSheet);

        secondSheet.setSelected(true);

        // 这里只能接收一个 index，base on 0
        wb.setActiveSheet(3);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
