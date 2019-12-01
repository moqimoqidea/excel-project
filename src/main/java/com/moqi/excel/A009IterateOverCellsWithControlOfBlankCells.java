package com.moqi.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.DEFAULT_PATH;

/**
 * 遍历单元格，控制丢失/空白的单元格
 * <p>
 * 在某些情况下，进行迭代时，您需要完全控制如何处理丢失或空白的行和单元格，
 * 并且需要确保访问每个单元格，而不仅仅是访问文件中定义的那些单元格。
 * （CellIterator将仅返回文件中定义的单元格，这在很大程度上是具有值或样式的单元格，但取决于Excel）。
 * <p>
 * 在这种情况下，应该获取一行的第一列和最后一列信息，
 * 然后调用getCell（int，MissingCellPolicy） 来获取单元格。
 * 使用 MissingCellPolicy 控制空白或空单元格的处理方式。
 *
 * @author moqi
 * On 12/1/19 11:20
 */
@Slf4j
public class A009IterateOverCellsWithControlOfBlankCells {

    private static final String A009_PATH = "A009.xlsx";

    public static void main(String[] args) throws InvalidFormatException, IOException {

        OPCPackage pkg = OPCPackage.open(new File(DEFAULT_PATH + A009_PATH));
        XSSFWorkbook workbook = new XSSFWorkbook(pkg);
        XSSFSheet sheet = workbook.getSheet("Sheet1");

        // 确定要处理的行
        int rowStart = Math.min(0, sheet.getFirstRowNum());
        int rowEnd = Math.max(5, sheet.getLastRowNum());
        for (int rowNum = rowStart; rowNum < rowEnd; rowNum++) {
            Row r = sheet.getRow(rowNum);
            if (r == null) {
                // 这整行是空的，根据需要进行处理
                log.info("第 {} 行是空的，跳过", rowNum);
                continue;
            }
            int lastColumn = Math.max(r.getLastCellNum(), 5);
            for (int cn = 0; cn < lastColumn; cn++) {
                Cell c = r.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);
                if (c == null) {
                    // 电子表格在此单元格中为空
                    log.info("第 {} 行，第 {} 列是空的，跳过", rowNum, cn);
                } else {
                    // 对单元格内容进行一些有用的操作
                    double numericCellValue = c.getNumericCellValue();
                    log.info("第 {} 行，第 {} 列不是空的，内容是：{}", rowNum, cn, numericCellValue);
                }
            }
        }

        pkg.close();


    }

}
