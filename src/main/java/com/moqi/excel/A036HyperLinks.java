package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.common.usermodel.HyperlinkType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

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
public class A036HyperLinks {


    private static final String HYPERLINKS_XLSX = "hyperlinks.xlsx";

    public static void main(String[] args) throws IOException {
        // 创建超链接
        createHyperLinks();

        // 读取超链接
        readHyperLinks();
    }

    private static void createHyperLinks() {

        Workbook wb = new XSSFWorkbook();
        CreationHelper createHelper = wb.getCreationHelper();

        // 超链接的单元格样式
        // 默认情况下，超链接为蓝色并带有下划线
        CellStyle cellStyle = wb.createCellStyle();
        Font font = wb.createFont();
        font.setUnderline(Font.U_SINGLE);
        font.setColor(IndexedColors.BLUE.getIndex());
        cellStyle.setFont(font);
        Cell cell;
        Sheet sheet = wb.createSheet("Hyperlinks");

        // URL
        cell = sheet.createRow(0).createCell(0);
        cell.setCellValue("URL Link");
        Hyperlink link = createHelper.createHyperlink(HyperlinkType.URL);
        link.setAddress("http://poi.apache.org/");
        cell.setHyperlink(link);
        cell.setCellStyle(cellStyle);

        // File
        cell = sheet.createRow(1).createCell(0);
        cell.setCellValue("File Link");
        link = createHelper.createHyperlink(HyperlinkType.FILE);
        // 同路径下只需要文件名
        link.setAddress("Book1.xlsx");
        cell.setHyperlink(link);
        cell.setCellStyle(cellStyle);

        // Email
        cell = sheet.createRow(2).createCell(0);
        cell.setCellValue("Email Link");
        link = createHelper.createHyperlink(HyperlinkType.EMAIL);

        // 注意，如果主题中包含空格，请确保其为url编码
        link.setAddress("mailto:poi@apache.org?subject=Hyperlinks");
        cell.setHyperlink(link);
        cell.setCellStyle(cellStyle);

        // 链接到此工作簿中的位置
        // 创建目标工作表和单元格
        Sheet sheet2 = wb.createSheet("Target Sheet");
        sheet2.createRow(0).createCell(0).setCellValue("Target Cell");
        cell = sheet.createRow(3).createCell(0);
        cell.setCellValue("Worksheet Link");
        Hyperlink link2 = createHelper.createHyperlink(HyperlinkType.DOCUMENT);
        link2.setAddress("'Target Sheet'!A1");
        cell.setHyperlink(link2);
        cell.setCellStyle(cellStyle);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + HYPERLINKS_XLSX);
    }

    private static void readHyperLinks() throws IOException {

        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLSX_DIR_PATH + HYPERLINKS_XLSX));
        Sheet sheet = wb.getSheet("Hyperlinks");

        Cell cell = sheet.getRow(0).getCell(0);
        Hyperlink link = cell.getHyperlink();
        if (link != null) {
            System.out.println(link.getAddress());
            log.info("Address:{}", link.getAddress());
        }

    }

}
