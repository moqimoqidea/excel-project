package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 将工作表设置为选中状态
 * 测试时打开 Excel 还是呈现 noSelectedSheet 表，等待再次测试
 *
 * @author moqi
 * On 12/1/19 19:49
 */
@Slf4j
public class A024SetSheetAsSelected {

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet noSelectedSheet = wb.createSheet("No Selected Sheet");
        Sheet selectedSheet = wb.createSheet("Selected Sheet");
        selectedSheet.setSelected(true);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
