package com.moqi.tool;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.FileOutputStream;

import static com.moqi.constant.Constant.YYYY_MM_DD;
import static com.moqi.constant.Constant.YYYY_MM_DD_HH_MM_SS;

/**
 * 项目公用工具类
 *
 * @author moqi
 * On 11/30/19 21:41
 */
@Slf4j
public class Tool {

    /**
     * 删除文件
     *
     * @param filePath 文件目录
     */
    public static void removeOldFile(String filePath) {
        try {
            File file = new File(filePath);

            boolean exists = file.exists();

            if (exists) {
                if (file.delete()) {
                    log.info("{} 文件已被删除", file.getName());
                } else {
                    log.info("文件删除失败");
                }
            } else {
                log.info("文件不存在无需删除");
            }

        } catch (Exception e) {
            log.warn("删除旧文件 方法内 发生异常");
        }
    }

    /**
     * 生成 Excel 文件
     *
     * @param workbook workbook
     * @param filePath 文件路径
     */
    public static void getExcelFile(Workbook workbook, String filePath) {
        try {
            Tool.removeOldFile(filePath);
            FileOutputStream out = new FileOutputStream(new File(filePath));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            log.warn("getExcelFile 方法发生异常");
        }
    }

    /**
     * 生成 YYYY_MM_DD 格式的 CellStyle
     *
     * @param workbook workbook
     * @return CellStyle
     */
    public static CellStyle getYmdCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(YYYY_MM_DD));

        return cellStyle;
    }

    /**
     * 生成 YYYY_MM_DD_HH_MM_SS 格式的 CellStyle
     *
     * @param workbook workbook
     * @return CellStyle
     */
    public static CellStyle getYmdHmsCellStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        CreationHelper createHelper = workbook.getCreationHelper();
        cellStyle.setDataFormat(createHelper.createDataFormat().getFormat(YYYY_MM_DD_HH_MM_SS));

        return cellStyle;
    }

}
