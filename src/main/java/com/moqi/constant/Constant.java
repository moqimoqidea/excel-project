package com.moqi.constant;

import java.text.SimpleDateFormat;

/**
 * 项目公用常量
 *
 * @author moqi
 * On 11/30/19 21:41
 */

public class Constant {

    /**
     * 测试默认使用 Excel 文件夹路径
     */
    public static final String DEFAULT_PATH = "src/main/resources/xlsx/";

    /**
     * 测试默认使用 Excel 文件路径
     */
    public static final String DEFAULT_EXCEL_PATH = "src/main/resources/xlsx/demo.xlsx";

    /**
     * 测试默认使用包内模板 Excel 文件地址
     */
    public static final String PACKAGE_SOURCE_PATH = "src/main/resources/xlsx/Book1.xlsx";

    /**
     * 字符串 YYYY_MM_DD_HH_MM_SS
     */
    public static final String YYYY_MM_DD_HH_MM_SS = "yyyy-MM-dd HH:mm:ss";

    /**
     * 字符串 YYYY_MM_DD
     */
    public static final String YYYY_MM_DD = "yyyy-MM-dd";

    /**
     * 日期格式化，格式为 YYYY_MM_DD_HH_MM_SS
     */
    public static final SimpleDateFormat SIMPLE_DATE_FORMAT_YYYY_MM_DD_HH_MM_SS = new SimpleDateFormat(YYYY_MM_DD_HH_MM_SS);

    /**
     * 日期格式化，格式为 YYYY_MM_DD
     */
    public static final SimpleDateFormat SIMPLE_DATE_FORMAT_YYYY_MM_DD = new SimpleDateFormat(YYYY_MM_DD);

}
