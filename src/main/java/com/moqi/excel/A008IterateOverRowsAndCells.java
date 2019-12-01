package com.moqi.excel;

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
 * 遍历行和单元格
 * 即使单纯使用 OPCPackage 读取 Excel 也会修改文件
 *
 * @author moqi
 * On 12/1/19 11:10
 */
@Slf4j
public class A008IterateOverRowsAndCells {

    public static void main(String[] args) throws InvalidFormatException, IOException {

        OPCPackage pkg = OPCPackage.open(new File(PACKAGE_SOURCE_PATH));
        XSSFWorkbook workbook = new XSSFWorkbook(pkg);

        for (Sheet sheet : workbook) {
            for (Row row : sheet) {
                for (Cell cell : row) {
                    String stringCellValue = cell.getStringCellValue();
                    log.info("stringCellValue:{}", stringCellValue);
                }
            }
        }

        pkg.close();

    }

}
