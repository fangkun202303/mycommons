package com.common.mycommons.util;

import cn.afterturn.easypoi.word.WordExportUtil;
import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.InputStream;
import java.util.Map;
import java.util.Objects;

/**
 * <p>
 * 导出word
 * </p>
 * Copyright (C) 2020 Kingstar Winning, Inc. All rights reserved.
 *
 * @author Kun.F Create on 2020/4/27 9:10
 * @version 1.0
 */
public class WordUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(WordUtil.class);

    /**
     * @description: word导出-easypoi导出
     *
     * @author: Fang Kun
     * @param fileName 模板的相对路径
     * @param exportName 导出时的名称
     * @param data 数据Map
     * @param response 响应
     * @date: 2020/4/27 9:16
     * @return: void
     */
    public static void wordExportTemplate(String fileName, String exportName, Map<String, Object> data, HttpServletResponse response){
        Objects.requireNonNull(fileName, "file path is not must be null");
        Objects.requireNonNull(exportName, "file name is not must be null");
        Objects.requireNonNull(data, "file data is not must be null");
        Objects.requireNonNull(response, "response is not must be null");
        InputStream resourceAsStream = Thread.currentThread().getContextClassLoader().getResourceAsStream(fileName);
        ServletOutputStream outputStream = null;
        try {
            MyXWPFDocument myXWPFDocument = new MyXWPFDocument(resourceAsStream);
            WordExportUtil.exportWord07(myXWPFDocument, data);
            PdfMulPageExportUtils.responseHandler(response, exportName);
            outputStream = response.getOutputStream();
            myXWPFDocument.write(outputStream);
            outputStream.close();
        }catch (Exception e){
            LOGGER.error(e.getMessage());
        }finally {
            FileUtiles.closeStreamOfOut(outputStream);
        }
    }
}
