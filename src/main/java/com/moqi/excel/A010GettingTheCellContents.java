package com.moqi.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;
import java.util.Date;

import static com.moqi.constant.Constant.DEFAULT_PATH;
import static com.moqi.constant.Constant.SIMPLE_DATE_FORMAT_YYYY_MM_DD;

/**
 * 获取单元格的内容
 * <p>
 * 要获取单元格的内容，您首先需要知道它是哪种单元格（例如，将字符串单元格作为其数字内容将获得NumberFormatException）。
 * 因此，您将需要打开单元格的类型，然后为该单元格调用适当的getter。
 * <p>
 * 在下面的代码中，我们遍历一张纸中的每个单元格，打印出单元格的引用（例如A3），然后打印出单元格的内容。
 *
 * @author moqi
 * On 12/1/19 14:44
 */
@Slf4j
public class A010GettingTheCellContents {

    private static final String A010_PATH = "A010.xlsx";

    public static void main(String[] args) throws IOException, InvalidFormatException {

        OPCPackage pkg = OPCPackage.open(new File(DEFAULT_PATH + A010_PATH));
        XSSFWorkbook workbook = new XSSFWorkbook(pkg);
        // 使用数据格式化类自动格式化
        DataFormatter formatter = new DataFormatter();
        // 通过 index 拿到 sheet
        Sheet sheet = workbook.getSheetAt(0);

        for (Row row : sheet) {
            for (Cell cell : row) {
                CellReference cellRef = new CellReference(row.getRowNum(), cell.getColumnIndex());
                String format = cellRef.formatAsString();
                // 通过获取单元格值并应用任何数据格式  (Date, 0.00, 1.23e9, $1.23, etc) 来获取显示在单元格中的文本
                String text = formatter.formatCellValue(cell);
                log.info("位置为 {} 的值自动格式化后的结果为 {}", format, text);
                // 或者，获取值并自行格式化
                switch (cell.getCellType()) {
                    case STRING:
                        String richString = cell.getRichStringCellValue().getString();
                        log.info("位置为 {} 的值类型为 String，值为 {}", format, richString);
                        break;
                    case NUMERIC:
                        if (DateUtil.isCellDateFormatted(cell)) {
                            Date dateCellValue = cell.getDateCellValue();
                            log.info("位置为 {} 的值类型为 日期类型，值为 {}", format, SIMPLE_DATE_FORMAT_YYYY_MM_DD.format(dateCellValue));
                        } else {
                            double numericCellValue = cell.getNumericCellValue();
                            log.info("位置为 {} 的值类型为 数字，值为 {}", format, numericCellValue);
                        }
                        break;
                    case BOOLEAN:
                        boolean booleanCellValue = cell.getBooleanCellValue();
                        log.info("位置为 {} 的值类型为 布尔值，值为 {}", format, booleanCellValue);
                        break;
                    case FORMULA:
                        String cellFormula = cell.getCellFormula();
                        log.info("位置为 {} 的值类型为 公式类型，值为 {}", format, cellFormula);
                        break;
                    case BLANK:
                        log.info("BLANK Cell");
                        log.info("位置为 {} 的值类型为 空白，值为 空白", format);
                        break;
                    default:
                        log.info("位置为 {} 的值类型没判断出来进入 Default 分支", format);
                }
            }
        }

        pkg.close();

    }

}
