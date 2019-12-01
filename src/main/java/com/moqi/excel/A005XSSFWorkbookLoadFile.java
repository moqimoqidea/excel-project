package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.PACKAGE_SOURCE_PATH;

/**
 * https://poi.apache.org/components/spreadsheet/quick-guide.html
 * <p>
 * If using HSSFWorkbook or XSSFWorkbook directly, you should generally go through POIFSFileSystem or OPCPackage,
 * to have full control of the lifecycle (including closing the file when done):
 * <p>
 * 如果直接使用 HSSFWorkbook 或 XSSFWorkbook，通常应通过 POIFSFileSystem 或 OPCPackage
 * 来完全控制生命周期（包括完成后关闭文件）：
 * <p>
 * 在来源和目标路径不一致的情况下，来源文件会变动（因而如果需要作为模板文件则需要解决这个问题）
 *
 * @author moqi
 * On 11/30/19 21:36
 */
@Slf4j
public class A005XSSFWorkbookLoadFile {

    private static final String TARGET_PATH = "/Users/moqi/Downloads/excel_test/Book2.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        OPCPackage pkg = OPCPackage.open(new File(PACKAGE_SOURCE_PATH));
        XSSFWorkbook workbook = new XSSFWorkbook(pkg);

        Sheet sheet1 = workbook.getSheet("Sheet1");
        Row row = sheet1.createRow(1);
        Cell cell = row.createCell(0);
        cell.setCellValue("OPCPackage");

        Tool.generateExcelFile(workbook, TARGET_PATH);
        pkg.close();

        log.info("程序执行完毕");
    }

}
