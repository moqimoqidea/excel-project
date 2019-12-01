package com.moqi.excel;

import com.moqi.tool.Tool;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 拆分和冻结窗格
 * 您可以创建两种类型的窗格：冻结窗格和拆分窗格。
 * <p>
 * 冻结窗格按列和行划分。使用以下机制创建冻结窗格：
 * <p>
 * sheet1.createFreezePane（3，2，3，2）;
 * <p>
 * 前两个参数是您希望分割的列和行。后两个参数指示在右下象限中可见的单元格。
 * <p>
 * 拆分窗格的显示方式有所不同。分割区域分为四个独立的工作区域。
 * 分割发生在像素级别，用户可以通过将其拖动到新位置来调整分割。
 * <p>
 * 通过以下调用创建拆分窗格：
 * <p>
 * sheet2.createSplitPane（2000，2000，0，0，Sheet.PANE_LOWER_LEFT）;
 * <p>
 * 第一个参数是拆分的x位置。这是一点的1/20。在这种情况下，一点似乎等于一个像素。
 * 第二个参数是分割的y位置。再次在1/20分之内。
 * <p>
 * 最后一个参数指示当前具有焦点的窗格。
 * 这将是Sheet.PANE_LOWER_LEFT，PANE_LOWER_RIGHT，PANE_UPPER_RIGHT或PANE_UPPER_LEFT之一。
 *
 * @author moqi
 * On 12/1/19 19:55
 */
@Slf4j
public class A026SplitsAndFreezePanes {


    public static void main(String[] args) {
        Workbook wb = new XSSFWorkbook();

        Sheet sheet1 = wb.createSheet("new sheet");
        Sheet sheet2 = wb.createSheet("second sheet");
        Sheet sheet3 = wb.createSheet("third sheet");
        Sheet sheet4 = wb.createSheet("fourth sheet");
        // 只冻结一行
        sheet1.createFreezePane(0, 1, 0, 1);
        // 只冻结一列
        sheet2.createFreezePane(1, 0, 1, 0);
        // 冻结列和行（忘记右下象限的滚动位置）。
        sheet3.createFreezePane(2, 2);
        // 创建一个拆分，其左下侧为活动象限
        sheet4.createSplitPane(2000, 2000, 0, 0, Sheet.PANE_LOWER_LEFT);

        Tool.generateExcelFile(wb, DEFAULT_XLSX_PATH);
    }

}
