package com.common.mycommons.util;

import com.alibaba.fastjson.JSON;
import freemarker.template.Template;

import java.util.Map;

/**
 *
 * @author Fang Kun$ Created on 2019/10/1016:35
 * @version 1.0
 */
public class PdfMulPageDTO{
    /**
     * 资源路径
     */
    private String path;

    /**
     * out fileName
     */
    private String outFileName;

    /**
     * 模板名称
     */
    private String templateName;

    /**
     * html 字符串
     */
    private Map<String,Object> dataHtml;

    /**
     * 页眉
     */
    private String pdfHeader;

    /**
     * 项目打包名称
     */
    private String itemName;

    /**
     * 模板对象
     */
    private Template template;

    /**
     * 字节码对象
     */
    private Class aClass;

    public PdfMulPageDTO() {
    }

    public PdfMulPageDTO(String path, String outFileName, Map<String,Object> dataHtml, String templateName, String pdfHeader, Template template, String itemName) {
        this.path=path;
        this.outFileName = outFileName;
        this.dataHtml = dataHtml;
        this.templateName = templateName;
        this.pdfHeader = pdfHeader;
        this.template=template;
        this.itemName=itemName;
    }

    public String getItemName() {
        return itemName;
    }

    public void setItemName(String itemName) {
        this.itemName = itemName;
    }

    public Class getaClass() {
        return aClass;
    }

    public void setaClass(Class aClass) {
        this.aClass = aClass;
    }

    public Template getTemplate() {
        return template;
    }

    public void setTemplate(Template template) {
        this.template = template;
    }

    public String getPath() {
        return path;
    }

    public void setPath(String path){
        this.path = path;
    }

    public String getOutFileName() {
        return outFileName;
    }

    public void setOutFileName(String outFileName) {
        this.outFileName = outFileName;
    }

    public String getTemplateName() {
        return templateName;
    }

    public void setTemplateName(String templateName) {
        this.templateName = templateName;
    }

    public Map<String, Object> getDataHtml() {
        return dataHtml;
    }

    public void setDataHtml(Map<String, Object> dataHtml) {
        this.dataHtml = dataHtml;
    }

    public String getPdfHeader() {
        return pdfHeader;
    }

    public void setPdfHeader(String pdfHeader) {
        this.pdfHeader = pdfHeader;
    }

    @Override
    public String toString() {
        return JSON.toJSONString(this);
    }
}
