package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 条件格式：个人理解类似格式刷
 * <p>
 * 在ConditionalFormats.java中查看有关Excel条件格式的更多示例
 * http://svn.apache.org/repos/asf/poi/trunk/src/examples/src/org/apache/poi/ss/examples/ConditionalFormats.java
 *
 * @author moqi
 * On 12/7/19 23:38
 */

@SuppressWarnings("AlibabaLowerCamelCaseVariableNaming")
public class A040ConditionalFormatting {

    public static void main(String[] args) {
        Workbook workbook = new HSSFWorkbook();
        Sheet sheet = workbook.createSheet();
        SheetConditionalFormatting sheetCF = sheet.getSheetConditionalFormatting();
        ConditionalFormattingRule rule1 = sheetCF.createConditionalFormattingRule(ComparisonOperator.EQUAL, "0");
        FontFormatting fontFmt = rule1.createFontFormatting();
        fontFmt.setFontStyle(true, false);
        fontFmt.setFontColorIndex(IndexedColors.DARK_RED.index);
        BorderFormatting bordFmt = rule1.createBorderFormatting();
        bordFmt.setBorderBottom(BorderStyle.THIN);
        bordFmt.setBorderTop(BorderStyle.THICK);
        bordFmt.setBorderLeft(BorderStyle.DASHED);
        bordFmt.setBorderRight(BorderStyle.DOTTED);
        PatternFormatting patternFmt = rule1.createPatternFormatting();
        patternFmt.setFillBackgroundColor(IndexedColors.YELLOW.index);
        ConditionalFormattingRule rule2 = sheetCF.createConditionalFormattingRule(ComparisonOperator.BETWEEN, "-10", "10");
        ConditionalFormattingRule[] cfRules =
                {
                        rule1, rule2
                };
        CellRangeAddress[] regions = {
                CellRangeAddress.valueOf("A3:A5")
        };
        sheetCF.addConditionalFormatting(regions, cfRules);

        Tool.generateExcelFile(workbook, DEFAULT_XLS_PATH);
    }

}
