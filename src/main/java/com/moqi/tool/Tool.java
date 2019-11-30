package com.moqi.tool;

import lombok.extern.slf4j.Slf4j;

import java.io.File;

/**
 * 项目公用工具类
 *
 * @author moqi
 * On 11/30/19 21:41
 */
@Slf4j
public class Tool {

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

}
