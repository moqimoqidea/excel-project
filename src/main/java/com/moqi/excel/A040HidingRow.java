package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 隐藏行和取消隐藏行
 * <p>
 * 使用Excel，可以通过以下方式隐藏工作表中的行：
 * 选择该行（一个或多个），右键单击鼠标右键，然后从出现的弹出菜单中选择“隐藏”。
 * <p>
 * 要使用POI对此进行仿真，只需在XSSFRow或HSSFRow的实例上调用setZeroHeight（）方法
 * （该方法在两个类都实现的ss.usermodel.Row接口上定义），如下所示：
 * <p>
 * Workbook workbook = new XSSFWorkbook();  // OR new HSSFWorkbook()
 * Sheet sheet = workbook.createSheet(0);
 * Row row = workbook.createRow(0);
 * row.setZeroHeight();
 * <p>
 * 如果文件现在已保存到光盘中，则第一页上的第一行将不可见。
 * <p>
 * 使用Excel，可以通过以下方法取消隐藏先前隐藏的行：
 * 选择要隐藏的行的上方和下方的行，然后按住Ctrl键，Shift和数字9并释放它们，然后再将其全部释放。
 * <p>
 * 要使用POI模拟此行为，请执行以下操作：
 * <p>
 * Workbook workbook = WorkbookFactory.create(new File(.......));
 * Sheet = workbook.getSheetAt(0);
 * Iterator<Row> row Iter = sheet.iterator();
 * while(rowIter.hasNext()) {
 * Row row = rowIter.next();
 * if(row.getZeroHeight()) {
 * row.setZeroHeight(false);
 * }
 * }
 * <p>
 * 如果现在将文件保存到磁盘，则工作簿第一页上以前隐藏的行现在将可见。
 * <p>
 * 该示例说明了两个功能。首先，可以通过调用setZeroHeight（）方法并传递布尔值'false'来取消隐藏行。
 * 其次，它说明了如何测试行是否被隐藏。只需调用getZeroHeight（）方法，
 * 如果该行被隐藏，它将返回“ true”，否则返回“ false”。
 *
 * @author moqi
 * On 12/7/19 23:43
 */
@Slf4j
public class A040HidingRow {

    public static void main(String[] args) throws IOException {
        // 隐藏
        // hideRow();

        // 展示（需要先运行隐藏然后注释掉隐藏）: 测试失败
        showRow();
    }

    private static void hideRow() {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Hiding Test");
        Row row = sheet.createRow(0);
        row.setZeroHeight(true);

        Tool.generateExcelFile(workbook, DEFAULT_XLSX_PATH);
    }

    private static void showRow() throws IOException {

        Workbook workbook = WorkbookFactory.create(new File(DEFAULT_XLSX_PATH));
        Sheet sheet = workbook.getSheet("Hiding Test");
        Row row = sheet.getRow(0);

        boolean zeroHeight = row.getZeroHeight();
        log.info("zeroHeight:{}", zeroHeight);
        row.setZeroHeight(false);

    }

}
