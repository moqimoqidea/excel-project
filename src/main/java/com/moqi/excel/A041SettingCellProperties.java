package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellUtil;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.HashMap;
import java.util.Map;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 设置单元格属性
 * 有时，创建具有基本样式的电子表格然后将特殊样式应用于某些单元格
 * （例如在一系列单元格周围绘制边框或设置区域填充）会更容易或更有效。
 * CellUtil.setCellProperties允许您执行此操作，而无需在电子表格中创建一堆不必要的中间样式。
 * <p>
 * 将属性创建为Map，并以以下方式将其应用于单元格。
 * <p>
 * 注意：这不会替换单元格的属性，它会将您放入Map中的属性与单元格的现有样式属性合并。
 * 如果属性已经存在，则将其替换为新属性。如果属性不存在，则将其添加。此方法不会删除CellStyle属性。
 *
 * @author moqi
 * On 12/7/19 23:43
 */
@SuppressWarnings("AlibabaUndefineMagicConstant")
@Slf4j
public class A041SettingCellProperties {

    public static void main(String[] args) {
        Workbook workbook = new XSSFWorkbook();
        Sheet sheet = workbook.createSheet("Sheet1");
        Map<String, Object> properties = new HashMap<>(32);

        // 单元格周围的边框
        properties.put(CellUtil.BORDER_TOP, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_BOTTOM, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_LEFT, BorderStyle.MEDIUM);
        properties.put(CellUtil.BORDER_RIGHT, BorderStyle.MEDIUM);

        // 给它加一个颜色（红色）
        properties.put(CellUtil.TOP_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.BOTTOM_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.LEFT_BORDER_COLOR, IndexedColors.RED.getIndex());
        properties.put(CellUtil.RIGHT_BORDER_COLOR, IndexedColors.RED.getIndex());

        // 将边界应用于B2处的单元格
        Row row = sheet.createRow(1);
        Cell cell = row.createCell(1);
        CellUtil.setCellStyleProperties(cell, properties);

        // 将边界应用于从D4开始的3x3区域
        for (int ix = 3; ix <= 5; ix++) {
            row = sheet.createRow(ix);
            for (int iy = 3; iy <= 5; iy++) {
                cell = row.createCell(iy);
                CellUtil.setCellStyleProperties(cell, properties);
            }
        }

        Tool.generateExcelFile(workbook, DEFAULT_XLSX_PATH);
    }


}
