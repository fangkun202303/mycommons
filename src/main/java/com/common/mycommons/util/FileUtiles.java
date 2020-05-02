package com.common.mycommons.util;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.validation.constraints.NotNull;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;

/**
 * <p>
 * 文件操作工具类
 * </p>
 *
 * @author Fang Kun Created on 2019/8/2714:27
 * @version 1.0
 */
public class FileUtiles {
    private static final Logger LOGGER = LoggerFactory.getLogger(FileUtiles.class);
    private static final String FILE_PREFIX="file:";
    private static final String FILE_SUFFIX=".jar";
    private static final String RESOURCE_PREFIX="template/";
    private static final String ITEM_NAME="hsa-ims-web";

    /**
     * @description 判断这个路径下面有没有文件
     *
     * @param filePath
     * @param fileNme
     * @return boolean
     * @author Fang Kun
     * @date 2019/10/25 12:23
     */
    public static boolean checkFile(@NotNull String filePath, @NotNull String fileNme){
        File file=new File(filePath);
        if(!file.isDirectory()){
            // 不是目录,就看这个路径的文件是不是和fileNme一样
            if(file.getName().equals(fileNme)){
                return false;
            }
        }else{
            // 是目录, 遍历目录下的文件,出去文件夹
            File[] files = file.listFiles();
            if(files==null){
                return false;
            }
            if(files.length<=0){
                return false;
            }
            for (File f :files){
                if(f!=null && f.getName().equals(fileNme)){
                    return true;
                }
            }
            return false;
        }
        return false;
    }

    /**
     * @description 再这个文件夹下面创建文件
     *
     * @return void
     * @author Fang Kun
     * @date 2019/10/25 12:24
     */
    public static void createFile(@NotNull String filePath, @NotNull String fileNme, @NotNull InputStream resourceAsStream){
        try (FileOutputStream fos = new FileOutputStream(filePath + fileNme)) {
            byte[] data = new byte[1024 * 8];
            int len = -1;
            while ((len = resourceAsStream.read(data)) != -1) {
                fos.write(data, 0, len);
            }
            closeStreamOfIn(resourceAsStream);
            closeStreamOfOut(fos);
        }catch (Exception e){
            LOGGER.error(e.getMessage());
        }
        LOGGER.debug( "createFile");
    }

    /**
     * @description: 获取文件资源路径
     *
     * @author: Fang Kun
     * @param resourcePath 资源路径
     * @param str
     * @param itemName
     * @date: 2020/3/18 13:26
     * @return: java.lang.String
     */
    public static String getFilePath(String resourcePath, String[] str, String itemName) {
        ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
        // 判断是开发工具启动还是jar运行
        if(resourcePath.indexOf(FILE_SUFFIX)>-1){
            // jar运行, 截取 项目名称.jar-hsa-ims-web.jar 之前的路径
            resourcePath=resourcePath.substring(0, resourcePath.indexOf(itemName));
        }else{
            resourcePath=resourcePath.substring(0, resourcePath.indexOf(ITEM_NAME));
        }
        LOGGER.debug("去掉协议之前: "+resourcePath);
        // 去掉 file: 协议
        if(resourcePath.indexOf(FILE_PREFIX)>-1){
            // 判断系统的类型,是windows 还是linux
//            String property = System.getProperty(ImsConstants.OS_NAME);
//            if(property.indexOf(ImsConstants.OS_NAME_FLAG_WIN)>-1){
//                resourcePath=resourcePath.substring(resourcePath.indexOf(FILE_PREFIX)+6);
//                // 说明是windows系统
//                resourcePath=resourcePath.replace(ImsConstants.OBLIQUE_LINE,ImsConstants.WINDOWS_SPLIT);
//            }else{
//                resourcePath=resourcePath.substring(resourcePath.indexOf(FILE_PREFIX)+5);
//            }
        }
        LOGGER.debug("去掉协议之后: "+resourcePath);
        // 判断这个路径下有没有这个文件,没有就创建
        for(String s : str){
            boolean b = checkFile(resourcePath, s);
            if(!b){
                // 说明这个目录下面没有这个文件
                InputStream inputStream = classLoader.getResourceAsStream(RESOURCE_PREFIX + s);
                createFile(resourcePath,s, inputStream);
            }
        }
        LOGGER.debug(resourcePath);
        return resourcePath;
    }

    /**
     * @description 关闭输出流
     *
     * @param outputStream
     * @return void
     * @author Fang Kun
     * @date 2019/10/30 20:07
     */
    @SuppressWarnings("all")
    public static void closeStreamOfOut(OutputStream outputStream){
        try {
            if(outputStream!=null){
                outputStream.close();
            }
        } catch (IOException e) {
            LOGGER.error(e.getMessage());
        }finally {
            if(outputStream!=null){
                try {
                    outputStream.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage());
                }
            }
        }
    }

    /**
     * @description 关闭输入流
     *
     * @param inputStream
     * @return void
     * @author Fang Kun
     * @date 2019/10/30 20:08
     */
    @SuppressWarnings("all")
    public static void closeStreamOfIn(InputStream inputStream){
        try {
            if(inputStream!=null){
                inputStream.close();
            }
        } catch (IOException e) {
            LOGGER.error(e.getMessage());
        }finally {
            if(inputStream!=null){
                try {
                    inputStream.close();
                } catch (IOException e) {
                    LOGGER.error(e.getMessage());
                }
            }
        }
    }

}