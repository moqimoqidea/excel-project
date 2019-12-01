package com.moqi.excel;

import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.extractor.ExcelExtractor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;

import static com.moqi.constant.Constant.DEFAULT_XLS_PATH;

/**
 * 文字提取
 * 对于大多数文本提取要求，标准的 ExcelExtractor 类应提供您所需要的全部。
 * <p>
 * 从 ExcelExtractor 得出只接受 HSSF 不接受 XSSF。
 *
 * @author moqi
 * On 12/1/19 15:14
 */
@Slf4j
public class A011TextExtraction {

    public static void main(String[] args) {
        try (InputStream inp = new FileInputStream(DEFAULT_XLS_PATH)) {
            HSSFWorkbook wb = new HSSFWorkbook(new POIFSFileSystem(inp));
            ExcelExtractor extractor = new ExcelExtractor(wb);
            // 是否不将计算公式的结果呈现
            extractor.setFormulasNotResults(true);
            // 是否包含 Sheet 名称
            extractor.setIncludeSheetNames(true);
            String text = extractor.getText();
            log.info("{}", text);
            wb.close();
        } catch (IOException e) {
            log.warn("A011TextExtraction.main 发生 IO 异常");
        }
    }

}
