package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.usermodel.*;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

/**
 * 数据验证
 * 从3.8版开始，POI的语法略有不同，可用于.xls和.xlsx格式的数据验证。
 * 本例以 xlsx 练习。
 * <p>
 * 创建基于XML的SpreadsheetML工作簿文件时，数据验证的工作原理类似；
 * 但是有区别。例如，在某些地方需要显式强制转换，因为xssf流中对数据验证的大部分支持已内置到统一的ss流中，其中更多。其他差异在代码中带有注释。
 *
 * @author moqi
 * On 12/7/19 22:19
 */

@SuppressWarnings("AlibabaRemoveCommentedCode")
public class A037DataValidations {

    private static final String A037_XLSX = "A037.xlsx";

    public static void main(String[] args) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet sheet = workbook.createSheet("Data Validation");


        // 根据一个或多个预定义值，检查用户输入到单元格中的值。
        // checkValueOnCell(sheet);

        // 下拉列表：该代码将执行相同的操作，但为用户提供一个下拉列表以从中选择一个值。
        // dropDownLists(sheet);
        // 错误消息：创建一个消息框，如果用户输入的值无效，该消息框将显示给用户。
        // 提示：创建提示，让用户看到包含数据验证的单元格何时获得焦点

        // 进一步的数据验证：
        // 要获得验证以检查输入的值是例如10到100之间的整数，
        // 请使用XSSFDataValidationHelper（s）createNumericConstraint（int，int，String，String）工厂方法。
        // checkNumberBetween10And100(sheet);
        // checkNumberBetween10And100(sheet) 方法验证了错误消息和提示

        // 传递给最后两个String参数的值可以是公式。“ =”符号用于表示公式。
        // 因此，以下将创建一个验证，仅当允许值落在两个单元格范围求和结果之间时
        // checkNumberFormula(sheet);
        // 如果调用createNumericConstraint（）方法是不可能创建下拉列表的，
        // 则setSuppressDropDownArrow（true）方法调用将被简单地忽略。
        // 请检查javadoc是否有其他约束类型，因为此处将不包括这些约束的示例。
        // 例如，在XSSFDataValidationHelper类上定义了一些方法，这些方法使您可以创建以下类型的约束。
        // 日期，时间，十进制，整数，数字，公式，文本长度和自定义约束。
        // 从电子表格单元格创建数据验证：
        // 上面未提及的另一种约束类型是公式列表约束。它允许您创建一个验证，并从一系列单元格中获取其值。这段代码
        // XSSFDataValidationConstraint dvConstraint =（XSSFDataValidationConstraint）
        //     dvHelper.createFormulaListConstraint（“ $ A $ 1：$ F $ 1”）;
        // 将创建一个验证，并从A1到F1范围的单元格中获取其值。
        // 如果使用这样的命名范围，则可以扩展该技术的实用性。
        checkValueFromCell(workbook);
        // 关于名称范围，OpenOffice Calc的规则略有不同。
        // Excel支持名称的Workbook和Sheet范围，但是Calc不支持，似乎仅支持名称的Sheet范围。
        // 因此，通常最好对这样的区域或区域名称进行完全限定。
        // name.setRefersToFormula("'Data Validation'!$B$1:$F$1");
        // 但是，这确实带来了另一个有趣的机会，那就是将所有用于验证的数据放入工作簿内隐藏工作表上的已命名单元格区域。
        // 然后可以在setRefersToFormula（）方法参数中明确标识这些范围。


        Tool.generateExcelFile(workbook, DEFAULT_XLSX_DIR_PATH + A037_XLSX);
    }

    /**
     * 可能因为原始输入值有问题测试未成功
     */
    private static void checkValueFromCell(XSSFWorkbook workbook) {
        XSSFName name = workbook.createName();
        XSSFSheet sheet = workbook.getSheet("Data Validation");
        name.setNameName("data");
        name.setRefersToFormula("$B$1:$F$1");
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                dvHelper.createFormulaListConstraint("data");
        CellRangeAddressList addressList = new CellRangeAddressList(
                0, 0, 0, 0);
        XSSFDataValidation validation = (XSSFDataValidation)
                dvHelper.createValidation(dvConstraint, addressList);
        validation.setSuppressDropDownArrow(true);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    /**
     * 都没有输入值的情况下，根据公式约束只能是 0 了
     */
    private static void checkNumberFormula(XSSFSheet sheet) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                dvHelper.createNumericConstraint(
                        XSSFDataValidationConstraint.ValidationType.INTEGER,
                        XSSFDataValidationConstraint.OperatorType.BETWEEN,
                        "=SUM(A1:A10)", "=SUM(B24:B27)");

        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
                dvConstraint, addressList);
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    private static void checkNumberBetween10And100(XSSFSheet sheet) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                dvHelper.createNumericConstraint(
                        XSSFDataValidationConstraint.ValidationType.INTEGER,
                        XSSFDataValidationConstraint.OperatorType.BETWEEN,
                        "10", "100");

        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
                dvConstraint, addressList);
        validation.setShowErrorBox(true);

        validation.setErrorStyle(DataValidation.ErrorStyle.STOP);
        validation.createErrorBox("输入参数错误", "参数必须介于 10-100 之间");

        validation.createPromptBox("提示：输入数字", "[10, 100]");
        validation.setShowPromptBox(true);

        sheet.addValidationData(validation);
    }

    /**
     * 限定第一个单元选择： 11、21、31
     *
     * @param sheet sheet
     */
    private static void dropDownLists(XSSFSheet sheet) {

        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});
        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
                dvConstraint, addressList);
        validation.setShowErrorBox(true);

        sheet.addValidationData(validation);

        // 请注意，可以简单地排除对setSuppressDropDownArrow（）方法的调用，或将其替换为：
        // 测试将此属性设置为 false 将导致下拉框消失
        validation.setSuppressDropDownArrow(true);
    }

    /**
     * 限定第一个单元格只能输入 11、21、31
     *
     * @param sheet sheet
     */
    private static void checkValueOnCell(XSSFSheet sheet) {
        XSSFDataValidationHelper dvHelper = new XSSFDataValidationHelper(sheet);
        XSSFDataValidationConstraint dvConstraint = (XSSFDataValidationConstraint)
                dvHelper.createExplicitListConstraint(new String[]{"11", "21", "31"});

        CellRangeAddressList addressList = new CellRangeAddressList(0, 0, 0, 0);
        XSSFDataValidation validation = (XSSFDataValidation) dvHelper.createValidation(
                dvConstraint, addressList);

        // 此处，布尔值false传递给setSuppressDropDownArrow（）方法。
        // 在上面的hssf.usermodel示例中，传递给此//方法的值为true。
        validation.setSuppressDropDownArrow(false);
        // 请注意此额外的方法调用。如果省略此方法调用，或者如果传递了布尔值false，
        // 则Excel将不会验证 用户输入到单元格中的值。
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

}
