package com.moqi.excel;

/**
 * 遍历单元格，控制丢失/空白的单元格
 * <p>
 * 在某些情况下，进行迭代时，您需要完全控制如何处理丢失或空白的行和单元格，并且需要确保访问每个单元格，而不仅仅是访问文件中定义的那些单元格。（CellIterator将仅返回文件中定义的单元格，这在很大程度上是具有值或样式的单元格，但取决于Excel）。
 * <p>
 * 在这种情况下，应该获取一行的第一列和最后一列信息，然后调用getCell（int，MissingCellPolicy） 来获取单元格。使用 MissingCellPolicy 控制空白或空单元格的处理方式。
 *
 * @author moqi
 * On 12/1/19 11:20
 */

public class A009IterateOverCellsWithControlOfBlankCells {

}
