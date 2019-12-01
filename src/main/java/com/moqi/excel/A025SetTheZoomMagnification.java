package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 设定变焦倍率
 * 缩放表示为分数。例如，要表示75％的缩放比例，请将3用作分子，将4用作分母。
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A025SetTheZoomMagnification {

    private static final int SEVENTY_FIVE = 75;

    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet zoomSheet = wb.createSheet("Zoom Sheet");
        // 调整缩放倍率到 75％
        zoomSheet.setZoom(SEVENTY_FIVE);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
