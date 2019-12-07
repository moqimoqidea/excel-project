package com.moqi.excel;

import com.moqi.tool.Tool;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;
import java.util.List;

import static com.moqi.constant.Constant.DEFAULT_XLSX_DIR_PATH;

/**
 * 图片
 * 图像是图形支持的一部分。要添加图像，只需在绘图族长上调用createPicture（）。
 * 在撰写本文时，支持以下类型：PNG、JPG、DIB
 * 应当注意，一旦将图像添加到图纸上，任何现有的图纸都可能被删除。
 * <p>
 * 警告
 * Picture.resize() 仅适用于JPEG和PNG。尚不支持其他格式。
 *
 * @author moqi
 * On 12/7/19 16:36
 */

public class A032Images {

    private static final String PICTURE_XLSX = "picture.xlsx";
    private static final String PICTURE_PATH = "src/main/resources/images/function.png";
    private static final String PICTURE_OUT_PATH = "src/main/resources/images/get_pict.png";
    private static final String PNG = "png";

    public static void main(String[] args) throws IOException {

        Workbook wb = new XSSFWorkbook();

        // 将图片数据添加到此工作簿。
        InputStream is = new FileInputStream(PICTURE_PATH);
        byte[] bytes = IOUtils.toByteArray(is);
        int pictureIdx = wb.addPicture(bytes, Workbook.PICTURE_TYPE_PNG);
        is.close();


        // 添加图片
        CreationHelper helper = wb.getCreationHelper();
        ClientAnchor anchor = helper.createClientAnchor();

        // 设置图片的左上角，
        anchor.setCol1(3);
        anchor.setRow1(2);

        Sheet sheet = wb.createSheet();
        // 创建绘图对象。这是所有形状的顶层容器。
        Drawing<?> drawing = sheet.createDrawingPatriarch();
        Picture pict = drawing.createPicture(anchor, pictureIdx);
        // 随后调用 Picture#resize() 将相对于它进行操作(相对于其左上角的图片自动调整大小)
        pict.resize();

        Tool.generateExcelFile(wb, DEFAULT_XLSX_DIR_PATH + PICTURE_XLSX);

        // 从工作簿中读取图像：
        List<? extends PictureData> allPictures = wb.getAllPictures();
        allPictures.forEach(x -> {
            String ext = x.suggestFileExtension();
            byte[] data = x.getData();

            if (PNG.equals(ext)) {
                try (OutputStream out = new FileOutputStream(PICTURE_OUT_PATH)) {
                    out.write(data);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        });

    }

}
