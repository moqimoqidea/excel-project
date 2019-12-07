package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 概述
 * 大纲非常适合将信息的各个部分组合在一起，并且可以使用POI API轻松地将其添加到列和行中。
 *
 * @author moqi
 * On 12/7/19 16:36
 */

public class A031Outlining {

    public static void main(String[] args) {

        Workbook wb = new HSSFWorkbook();
        Sheet sheet1 = wb.createSheet("new sheet");
        sheet1.groupRow(5, 14);
        sheet1.groupRow(7, 14);
        sheet1.groupRow(16, 19);
        sheet1.groupColumn(4, 7);
        sheet1.groupColumn(9, 12);
        sheet1.groupColumn(10, 11);

        // 要折叠（或展开）轮廓，请使用以下调用。您选择的行/列应包含一个已创建的组。它可以在组中的任何位置。
        sheet1.setRowGroupCollapsed(7, true);
        sheet1.setColumnGroupCollapsed(4, true);


        Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);
    }

}
