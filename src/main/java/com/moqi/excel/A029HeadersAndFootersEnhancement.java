package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Header;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFHeaderFooterProperties;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 页眉和页脚的 XSS 加强，区分首页和奇偶页
 * 当删除首页和偶数页的配置时，所有页保持一致。
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A029HeadersAndFootersEnhancement {


    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();
        XSSFSheet sheet = (XSSFSheet) wb.createSheet("new sheet");

        Tool.fill10000ValueOnSheet(sheet);

        // 创建首页标题
        Header header = sheet.getFirstHeader();
        header.setCenter("Center First Page Header");
        header.setLeft("Left First Page Header");
        header.setRight("Right First Page Header");

        // 创建一个偶数页眉
        Header header2 = sheet.getEvenHeader();
        header2.setCenter("Center Even Page Header");
        header2.setLeft("Left Even Page Header");
        header2.setRight("Right Even Page Header");

        // 创建一个奇数页头
        Header header3 = sheet.getOddHeader();
        header3.setCenter("Center Odd Page Header");
        header3.setLeft("Left Odd Page Header");
        header3.setRight("Right Odd Page Header");

        // 设置/删除标题属性
        XSSFHeaderFooterProperties prop = sheet.getHeaderFooterProperties();
        prop.setAlignWithMargins(true);
        prop.setScaleWithDoc(true);
        // 这会删除首页的页眉或页脚
        prop.removeDifferentFirst();
        // 这会删除偶数页的页眉或页脚
        prop.removeDifferentOddEven();

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }


}
