package com.moqi.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackagePart;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xslf.usermodel.XSLFSlideShow;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.xmlbeans.XmlException;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;

import static com.moqi.constant.Constant.DEFAULT_XLSX_PATH;

/**
 * 嵌入式对象
 * 可以对嵌入式Excel，Word或PowerPoint文档执行更详细的处理，或者与任何其他类型的嵌入式对象一起使用。
 * 以 xssf 为例
 * (Since POI-3.7)
 * <p>
 * 读取所有的嵌入式对象
 *
 * @author moqi
 * On 12/7/19 23:10
 */

@SuppressWarnings("StatementWithEmptyBody")
@Slf4j
public class A038EmbeddedObjects {

    public static void main(String[] args) throws IOException, OpenXML4JException, XmlException {
        XSSFWorkbook workbook = (XSSFWorkbook) WorkbookFactory.create(new File(DEFAULT_XLSX_PATH));
        for (PackagePart pPart : workbook.getAllEmbeddedParts()) {
            String contentType = pPart.getContentType();
            // Excel工作簿-二进制或OpenXML
            switch (contentType) {
                case "application/vnd.ms-excel": {
                    HSSFWorkbook embeddedWorkbook = new HSSFWorkbook(pPart.getInputStream());
                    break;
                }
                // Excel工作簿-OpenXML文件格式
                case "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet": {
                    OPCPackage docPackage = OPCPackage.open(pPart.getInputStream());
                    XSSFWorkbook embeddedWorkbook = new XSSFWorkbook(docPackage);
                    break;
                }
                // Word文档-二进制（OLE2CDF）文件格式
                case "application/msword":
                    // HWPFDocument document = new HWPFDocument(pPart.getInputStream());
                    break;
                // Word文档-OpenXML文件格式
                case "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {
                    OPCPackage docPackage = OPCPackage.open(pPart.getInputStream());
                    XWPFDocument document = new XWPFDocument(docPackage);
                    break;
                }
                // PowerPoint文档-二进制文件格式
                case "application/vnd.ms-powerpoint":
                    // HSLFSlideShow slideShow = new HSSFSlideShow(pPart.getInputStream());
                    break;
                // PowerPoint文档-OpenXML文件格式
                case "application/vnd.openxmlformats-officedocument.presentationml.presentation": {
                    OPCPackage docPackage = OPCPackage.open(pPart.getInputStream());
                    XSLFSlideShow slideShow = new XSLFSlideShow(docPackage);
                    break;
                }
                // 任何其他类型的嵌入式对象。
                default:
                    log.info("Unknown Embedded Document: {}", contentType);
                    InputStream inputStream = pPart.getInputStream();
                    break;
            }
        }
    }

}
