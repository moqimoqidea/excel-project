package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.AreaReference;
import org.apache.poi.ss.util.CellReference;

import java.io.File;
import java.io.IOException;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 命名范围是一种通过名称引用一组单元格的方法。命名单元格是命名范围的简并的情况，因为“单元格组”恰好包含一个单元格。
 * 您可以在工作簿中按其命名范围创建和引用单元格。
 * 使用命名范围时，将使用org.apache.poi.ss.util.CellReference 和org.apache.poi.ss.util.AreaReference类。
 * <p>
 * 注意：使用相对值（例如'A1：B1'）会导致在Microsoft Excel中使用工作簿时名称所指向的单元格的意外移动，
 * 通常使用'$ A $ 1：$ B $ 1'这样的绝对引用可以避免这种情况，请参阅下面连接
 * https://superuser.com/questions/800694/named-ranges-changing-randomly-in-excel-2010/1031047#1031047。
 * <p>
 * 创建命名范围/命名单元格
 *
 * @author moqi
 * On 12/7/19 16:36
 */

public class A033NamedRangesAndNamedCells {

    public static void main(String[] args) throws IOException {
        // 创建命名范围/命名单元格
        creatingNamedRangeAndNamedCell();
        // 从命名范围/命名单元格读取
        readingFromNamedRangeAndNamedCell();
        // 从不连续的命名范围读取
        readingFromNonContiguousNamedRanges();
        // 当心
        beware();
    }

    private static void creatingNamedRangeAndNamedCell() {
        // 创建命名范围/命名单元格
        String sname = "TestSheet", cname = "TestName", cvalue = "TestVal";
        Workbook wb = new HSSFWorkbook();
        Sheet sheet = wb.createSheet(sname);
        sheet.createRow(0).createCell(0).setCellValue(cvalue);

        // 1. 使用 areaReference 为单个单元格创建命名范围
        Name namedCell = wb.createName();
        namedCell.setNameName(cname + "1");
        // area reference
        String reference = sname + "!$A$1:$A$1";
        namedCell.setRefersToFormula(reference);

        // 2. 使用 cellReference 为单个单元格创建命名范围
        Name namedCel2 = wb.createName();
        namedCel2.setNameName(cname + "2");
        // cell reference
        reference = sname + "!$A$1";
        namedCel2.setRefersToFormula(reference);

        // 3. 使用 AreaReference为 区域创建命名范围
        Name namedCel3 = wb.createName();
        namedCel3.setNameName(cname + "3");
        // area reference
        reference = sname + "!$A$1:$C$5";
        namedCel3.setRefersToFormula(reference);

        // 4. 创建命名公式
        Name namedCel4 = wb.createName();
        namedCel4.setNameName("my_sum");
        namedCel4.setRefersToFormula("SUM(" + sname + "!$I$2:$I$6)");

        Tool.generateExcelFile(wb, DEFAULT_XLS_PATH);
    }

    private static void readingFromNamedRangeAndNamedCell() throws IOException {
        String cname = "TestName";
        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLS_PATH));

        // 检索命名范围
        int namedCellIdx = wb.getNameIndex(cname);
        Name aNamedCell = wb.getNameAt(namedCellIdx);

        // 检索指定范围内的单元格并测试其内容
        AreaReference aref = new AreaReference(aNamedCell.getRefersToFormula(), null);
        CellReference[] crefs = aref.getAllReferencedCells();
        for (CellReference cref : crefs) {
            Sheet sheet = wb.getSheet(cref.getSheetName());
            Row row = sheet.getRow(cref.getRow());
            //noinspection unused
            Cell cell = row.getCell(cref.getCol());
            // 根据单元格类型等提取单元格内容
        }
    }

    private static void readingFromNonContiguousNamedRanges() throws IOException {
        String cname = "TestName";
        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLS_PATH));

        // 检索命名范围
        // 将类似于 "$C$10,$D$12:$D$14";
        int namedCellIdx = wb.getNameIndex(cname);
        Name aNamedCell = wb.getNameAt(namedCellIdx);

        // 检索指定范围内的单元格并测试其内容
        // 将获得C10的一个AreaReference，然后
        //  D12到D14的另一个
        AreaReference[] arefs = AreaReference.generateContiguous(null, aNamedCell.getRefersToFormula());
        for (AreaReference aref : arefs) {
            // 只获得区域的角落
            // （使用arefs [i] .getAllReferencedCells（）获取所有单元格）
            CellReference[] crefs = aref.getAllReferencedCells();
            for (CellReference cref : crefs) {
                // 检查它变成真实的东西
                Sheet sheet = wb.getSheet(cref.getSheetName());
                Row row = sheet.getRow(cref.getRow());
                //noinspection unused
                Cell cell = row.getCell(cref.getCol());
                // 对这个角落单元格进行操作
            }
        }
    }

    /**
     * 请注意，删除单元格后，Excel不会删除附加的命名范围。
     * 因此，工作簿可以包含指向不再存在的单元格的命名范围。
     * 在构造AreaReference之前，您应该检查参考的有效性。
     */
    private static void beware() throws IOException {
        String cname = "TestName";
        Workbook wb = WorkbookFactory.create(new File(DEFAULT_XLS_PATH));
        int namedCellIdx = wb.getNameIndex(cname);
        Name aNamedCell = wb.getNameAt(namedCellIdx);

        //noinspection StatementWithEmptyBody
        if (aNamedCell.isDeleted()) {
            // 已命名的范围指向已删除的单元格。
        } else {
            //noinspection unused
            AreaReference ref = new AreaReference(aNamedCell.getRefersToFormula(), null);
        }
    }

}
