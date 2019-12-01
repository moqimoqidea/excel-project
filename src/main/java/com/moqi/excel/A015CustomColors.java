package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.*;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;
import static com.moqi.constant.Constant.DEFAULT_XLS_DIR_PATH;

/**
 * Excel 使用用户自定义的颜色
 *
 * @author moqi
 * On 12/1/19 16:39
 */

public class A015CustomColors {

    private static final String DEFAULT_PALETTE_XLS = "default_palette.xls";
    private static final String MODIFIED_PALETTE_XLS = "modified_palette.xls";
    private static final String DEFAULT_PALETTE = "Default Palette";
    private static final String MODIFIED_PALETTE = "Modified Palette";
    private static final String CUSTOM_COLOR_XLSX = "custom_color.xlsx";

    public static void main(String[] args) {

        customHssfColor();

        customXssfColor();
    }

    /**
     * 自定义 xls 文件单元格颜色
     */
    private static void customHssfColor() {
        HSSFWorkbook wb = new HSSFWorkbook();
        HSSFSheet sheet = wb.createSheet();
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = row.createCell(0);
        cell.setCellValue(DEFAULT_PALETTE);

        /*
         * 应用标准调色板中的某些颜色，与前面的示例相同:
         * 我们将在石灰背景上使用红色文本
         */
        HSSFCellStyle style = wb.createCellStyle();
        style.setFillForegroundColor(HSSFColor.HSSFColorPredefined.LIME.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        HSSFFont font = wb.createFont();
        font.setColor(HSSFColor.HSSFColorPredefined.RED.getIndex());
        style.setFont(font);
        cell.setCellStyle(style);
        // 使用默认调色板保存
        Tool.generateExcelFile(wb, DEFAULT_XLS_DIR_PATH + DEFAULT_PALETTE_XLS);

        /*
         * 现在，让我们替换调色板中的RED和LIME
         * 具有更有吸引力的组合
         * （很喜欢从freebsd.org借来的）
         */
        cell.setCellValue(MODIFIED_PALETTE);
        // 为工作簿创建自定义调色板
        HSSFPalette palette = wb.getCustomPalette();
        // 用 freebsd.org red替换标准red
        palette.setColorAtIndex(HSSFColor.HSSFColorPredefined.RED.getIndex(),
                //RGB red (0-255)
                (byte) 153,
                //RGB green
                (byte) 0,
                //RGB blue
                (byte) 0
        );

        // 用 freebsd.org 黄金代替石灰
        palette.setColorAtIndex(HSSFColor.HSSFColorPredefined.LIME.getIndex(),
                (byte) 255,
                (byte) 204,
                (byte) 102
        );

        /*
         * 保存修改后的调色板
         * 请注意，无论我们以前使用过RED还是LIME，新颜色神奇地出现
         */
        Tool.generateExcelFile(wb, DEFAULT_XLS_DIR_PATH + MODIFIED_PALETTE_XLS);
    }

    /**
     * 自定义 xlsx 文件单元格颜色
     */
    private static void customXssfColor() {
        XSSFWorkbook wb = new XSSFWorkbook();
        XSSFSheet sheet = wb.createSheet();
        XSSFRow row = sheet.createRow(0);
        XSSFCell cell = row.createCell(0);
        cell.setCellValue("custom XSSF colors");
        XSSFCellStyle style1 = wb.createCellStyle();

        style1.setFillForegroundColor(new XSSFColor(
                new java.awt.Color(128, 100, 0),
                new DefaultIndexedColorMap()));

        style1.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        cell.setCellStyle(style1);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + CUSTOM_COLOR_XLSX);
    }

}
