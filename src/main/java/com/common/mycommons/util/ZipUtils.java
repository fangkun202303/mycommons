package com.common.mycommons.util;

import org.apache.commons.io.IOUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.List;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

/**
 * <P>
 * 描述：zip文件工具类，压缩，解压
 * </p>
 *
 * @author lishang Created on 2020/4/2110:57
 * @version 1.0
 */
public class ZipUtils {

    private static final Logger LOGGER = LoggerFactory.getLogger(ZipUtils.class);

    /**
     * 创建ZipOutPutStream
     *
     * @param outputStream
     * @return
     */
    public static ZipOutputStream createZipOutputStream(OutputStream outputStream) {
        ZipOutputStream zipOutputStream = new ZipOutputStream(outputStream);
        return zipOutputStream;
    }

    /**
     * 将字符串写入到zip流
     *
     * @param content
     * @param storeFullFileName
     * @param zipOutputStream
     * @throws IOException
     */
    public static void doCompress(String content, String storeFullFileName, ZipOutputStream zipOutputStream) {
        if (content == null) {
            return;
        }
        try {
            zipOutputStream.putNextEntry(new ZipEntry(storeFullFileName));
            IOUtils.write(content, zipOutputStream, StandardCharsets.UTF_8.name());
        } catch (IOException e) {
            LOGGER.error("doCompress error:{}", e.getMessage());
        } finally {
            try {
                zipOutputStream.closeEntry();
            } catch (IOException e) {
                LOGGER.error("doCompress closeEntry error :{}", e.getMessage());
            }
        }
    }

    /**
     * 将输出流写入到zip流
     *
     * @param inputStream
     * @param storeFullFileName
     * @param zipOutputStream
     * @throws IOException
     */
    public static void doCompress(InputStream inputStream, String storeFullFileName, ZipOutputStream zipOutputStream) {
        if (inputStream == null) {
            return;
        }
        try {
            zipOutputStream.putNextEntry(new ZipEntry(storeFullFileName));
            byte[] bytes = IOUtils.toByteArray(inputStream);
            zipOutputStream.write(bytes);
        } catch (IOException e) {
            LOGGER.error("doCompress error:{}", e.getMessage());
        } finally {
            try {
                zipOutputStream.closeEntry();
            } catch (IOException e) {
                LOGGER.error("doCompress closeEntry error :{}", e.getMessage());
            }
        }
    }

    /**
     * 将多个文件写入Zip流中
     *
     * @param srcFiles
     * @param storeFileDir
     * @param zipOutputStream
     */
    public static void doCompressFiles(List<File> srcFiles, String storeFileDir, ZipOutputStream zipOutputStream) throws IOException {
        for (File srcFile : srcFiles) {
            String fileName = srcFile.getName();
            try {
                zipOutputStream.putNextEntry(new ZipEntry(storeFileDir + File.separator + fileName));
                FileInputStream fileInputStream = new FileInputStream(srcFile);
                byte[] bytes = IOUtils.toByteArray(fileInputStream);
                zipOutputStream.write(bytes);
                fileInputStream.close();
            } catch (IOException e) {
                LOGGER.error("doCompressFiles  error :{}", e.getMessage());
            } finally {
                try {
                    zipOutputStream.closeEntry();
                } catch (IOException e) {
                    LOGGER.error("doCompressFiles closeEntry error :{}", e.getMessage());
                }
            }
        }
    }

    /**
     * 解压zip文件
     *
     * @param zipInputStream
     * @param zipRootName
     * @return 解压后的文件夹路径
     * @throws IOException
     */
    public static String doDeCompressZipFile(ZipInputStream zipInputStream, String zipRootName) {
        ZipEntry nextEntry = null;
        String tempPath = System.getProperty("java.io.tmpdir");
        try {
            while (((nextEntry = zipInputStream.getNextEntry()) != null) && !nextEntry.isDirectory()) {
                // 如果entry不为空，并不在同一目录下， 获取文件目录
                File file = new File(tempPath + File.separator + nextEntry.getName());
                boolean newFile=false;
                // 如果该文件不存在
                if (!file.exists()) {
                    File fileParent = file.getParentFile();
                    if (!fileParent.exists()) {
                        fileParent.mkdirs();
                    }
                    // 创建文件
                    newFile = file.createNewFile();
                }
                byte[] bytes = IOUtils.toByteArray(zipInputStream);
                if (newFile){
                    FileOutputStream fileOutputStream = new FileOutputStream(file);
                    IOUtils.write(bytes, fileOutputStream);
                    fileOutputStream.close();
                }
                // 关闭当前entry
                zipInputStream.closeEntry();
            }
        } catch (IOException e) {
            LOGGER.error("doDeCompressZipFile. error :{}", e.getMessage());
        } finally {
            try {
                // 关闭流
                zipInputStream.close();
            } catch (IOException e) {
                LOGGER.error("zipInputStream.close error :{}", e.getMessage());
            }
        }
        StringBuilder sb = new StringBuilder(tempPath).append(File.separator).append(zipRootName);
        return sb.toString();
    }

}
