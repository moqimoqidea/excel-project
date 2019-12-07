package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PropertyTemplate;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 绘制边框
 * 在Excel中，只需按一下按钮，便可以在整个工作簿区域上应用一组边框。
 * PropertyTemplate对象使用定义为允许在单元格区域周围绘制顶部，底部，左侧，左侧，右侧，
 * 水平，垂直，内部，外部或所有边界的方法和常量来模拟此情况。
 * 其他方法允许将颜色应用于边框。
 * <p>
 * 它的工作原理是这样的：创建一个PropertyTemplate对象，该对象是要应用于图纸的边框的容器。
 * 然后，将边框和颜色添加到PropertyTemplate，最后将其应用于需要该边框集的任何图纸上。
 * 您可以创建多个PropertyTemplate对象，然后将它们应用于单个图纸，
 * 也可以将同一PropertyTemplate对象应用于多个图纸。就像预印表格一样。
 * <p>
 * 枚举：
 * <p>
 * 边框样式
 * 定义边框的外观，粗或细，实线或虚线，单或双。
 * 该枚举替换了已弃用的CellStyle.BORDER_XXXXX常量。
 * PropertyTemplate将不支持较早样式的BORDER_XXXXX常量。
 * 一个特殊的BorderStyle.NONE值将在应用单元格后将其从单元格中删除。
 * 边界范围
 * 描述边框样式将应用于的区域部分。例如，“顶部”，“底部”，“内部”或“外部”。
 * BorderExtent.NONE的特殊值将从PropertyTemplate中删除边框。
 * 应用模板时，不会对PropertyTemplate中不存在边框属性的单元格边框进行任何更改。
 *
 * 注意：最后一个pt.drawBorders（）调用使用BorderStyle.NONE从范围中删除边框。
 * 像setCellStyleProperties一样，applyBorders方法合并单元格样式的属性，
 * 因此仅当现有边框被其他东西替换时，才更改现有边框，或者仅当边框被BorderStyle.NONE替换时才删除现有边框。
 * 若要从边框删除颜色，请使用IndexedColor.AUTOMATIC.getIndex（）。
 *
 * 此外，要从PropertyTemplate对象中删除边框或颜色，请使用BorderExtent.NONE。
 * 目前尚不适用于对角线边框。
 *
 * @author moqi
 * On 12/7/19 23:43
 */
@Slf4j
public class A042DrawingBorders {

    public static void main(String[] args) {


        // 绘制边框（三个3x3网格）
        PropertyTemplate pt = new PropertyTemplate();
        // ＃1）这些边框的默认颜色均为中等
        pt.drawBorders(new CellRangeAddress(1, 3, 1, 3),
                BorderStyle.MEDIUM, BorderExtent.ALL);
        // ＃2）这些单元格将具有中等的外部边界和较薄的内部边界
        pt.drawBorders(new CellRangeAddress(5, 7, 1, 3),
                BorderStyle.MEDIUM, BorderExtent.OUTSIDE);
        pt.drawBorders(new CellRangeAddress(5, 7, 1, 3), BorderStyle.THIN,
                BorderExtent.INSIDE);

        // ＃3）这些单元格将全部为中等重量，且颜色不同外部，内部水平和内部垂直边界。中心单元格将没有边框。
        pt.drawBorders(new CellRangeAddress(9, 11, 1, 3),
                BorderStyle.MEDIUM, IndexedColors.RED.getIndex(),
                BorderExtent.OUTSIDE);
        pt.drawBorders(new CellRangeAddress(9, 11, 1, 3),
                BorderStyle.MEDIUM, IndexedColors.BLUE.getIndex(),
                BorderExtent.INSIDE_VERTICAL);
        pt.drawBorders(new CellRangeAddress(9, 11, 1, 3),
                BorderStyle.MEDIUM, IndexedColors.GREEN.getIndex(),
                BorderExtent.INSIDE_HORIZONTAL);
        pt.drawBorders(new CellRangeAddress(10, 10, 2, 2),
                BorderStyle.NONE,
                BorderExtent.ALL);

        // 在工作表上应用边框
        Workbook wb = new XSSFWorkbook();
        Sheet sh = wb.createSheet("Sheet1");
        pt.applyBorders(sh);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }


}
