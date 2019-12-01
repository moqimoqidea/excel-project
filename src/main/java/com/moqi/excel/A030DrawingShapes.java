package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFClientAnchor;
import org.apache.poi.hssf.usermodel.HSSFPatriarch;
import org.apache.poi.hssf.usermodel.HSSFSimpleShape;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 绘图形状：
 * POI支持使用Microsoft Office绘图工具绘制图形。图纸上的形状按组和形状的层次结构进行组织。
 * 最重要的形状是族长。这一点在工作表上根本看不到。
 * 要开始绘制,您需要 在HSSFSheet类上调用createPatriarch。
 * 这具有擦除存储在该纸张中的任何其他形状信息的效果。
 * 默认情况下,除非您调用此方法,否则POI会将形状记录留在工作表中。
 * <p>
 * 要创建形状,您必须执行以下步骤：
 * 创建族长。
 * 创建锚点以将形状放置在图纸上。
 * 要求族长创造形状。
 * 设置形状类型（线,椭圆,矩形等）
 * 设置有关形状的其他样式详细信息。（例如：线宽等）
 * <p>
 * 本例使用 xls 测试，画了一条斜线。
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A030DrawingShapes {


    public static void main(String[] args) {
        HSSFWorkbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet("New Sheet");

        HSSFPatriarch patriarch = (HSSFPatriarch) sheet.createDrawingPatriarch();
        HSSFClientAnchor anchor = new HSSFClientAnchor(0, 0, 1023, 255, (short) 1, 0, (short) 1, 0);
        HSSFSimpleShape shape1 = patriarch.createSimpleShape(anchor);
        shape1.setShapeType(HSSFSimpleShape.OBJECT_TYPE_LINE);

        Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);
    }


}
