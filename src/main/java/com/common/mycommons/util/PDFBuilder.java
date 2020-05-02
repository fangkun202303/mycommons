package com.common.mycommons.util;

import cn.hutool.core.util.ObjectUtil;
import com.alibaba.fastjson.JSON;
import com.itextpdf.text.*;
import com.itextpdf.text.pdf.BaseFont;
import com.itextpdf.text.pdf.ColumnText;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfPageEventHelper;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import org.apache.commons.lang3.StringUtils;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;

/**
 * <p>
 * 设置页面附加属性
 * </p>
 *
 * @author Fang Kun Created on 2019/8/2714:27
 * @version 1.0
 */
public class PDFBuilder extends PdfPageEventHelper {
    private static final Logger LOGGER = LoggerFactory.getLogger(PDFBuilder.class);
    /**
     * 字体文件名
     */
    private String fontFileName;
    /**
     * 基础字体对象
     */
    private BaseFont bf;
    /**
     * 利用基础字体生成的字体对象，一般用于生成中文文字
     */
    private Font fontDetail;
    /**
     * 文档字体大小
     */
    private  int fontSize=12;
    /**
     * 模板
     */
    private PdfTemplate template;
    /**
     * 数据实体
     */
    private PdfMulPageDTO pdfMulPageDTO;

    /**
     * 不允许空的构造方法
     */
    private PDFBuilder() {

    }

    /**
     * 初始化字体,与数据
     */
    public PDFBuilder( PdfMulPageDTO pdfMulPageDTO) {
        this.pdfMulPageDTO=pdfMulPageDTO;
        this.fontFileName="ping_fang_light.ttf";
    }

    public int getPresentFontSize() {
        return fontSize;
    }

    public void setPresentFontSize(int fontSize) {
        this.fontSize = fontSize;
    }

    public String getFontFileName() {
        return fontFileName;
    }

    public void setFontFileName(String fontFileName) {
        this.fontFileName = fontFileName;
    }

    /**
     * @description: 文档打开时创建模板
     *
     * @author: Fang Kun
     * @param writer
     * @param document
     * @date: 2020/3/18 13:28
     * @return: void
     */
    @Override
    public void onOpenDocument(PdfWriter writer, Document document) {
        template = writer.getDirectContent().createTemplate(50, 50);
    }

    /**
     * @description: 关闭每页的时候，写入页眉,页脚等
     *
     * @author: Fang Kun
     * @param writer
     * @param document
     * @date: 2020/3/18 13:27
     * @return: void
     */
    @Override
    public void onEndPage(PdfWriter writer, Document document) {
        if(pdfMulPageDTO==null){
            return ;
        }

        initFront();

        try {
            if (fontDetail == null) {
                // 数据字体
                fontDetail = new Font(bf, fontSize, Font.NORMAL);
            }
        } catch (Exception e) {
            e.printStackTrace();
        }

        // 1.写入页眉
        ColumnText.showTextAligned(
                writer.getDirectContent(),
                Element.ALIGN_LEFT,
                new Phrase(pdfMulPageDTO.getPdfHeader(), fontDetail),
                document.left(),
                document.top() + 25, 0);


        // 2.写入前半部分的 第 X页/共
        int pageS = writer.getPageNumber();
        String foot1 = "第 " + pageS + " 页 / 共";
        Phrase footer = new Phrase(foot1, fontDetail);
        // 3.计算前半部分的foot1的长度，后面好定位最后一部分的'Y页'这俩字的x轴坐标，字体长度也要计算进去 = len
        float len = bf.getWidthPoint(foot1, fontSize);
        // 4.拿到当前的PdfContentByte
        PdfContentByte cb = writer.getDirectContent();
        // 5.写入页脚1，x轴就是(右margin+左margin + right() -left()- len)/2.0F
        // ,y轴就是底边界-25,否则就贴边重叠到数据体里了就不是页脚了；注意Y轴是从下往上累加的，最上方的Top值是大于Bottom好几百开外的。
        ColumnText.showTextAligned(cb, Element.ALIGN_RIGHT, footer,(document.rightMargin() + document.right() + document.leftMargin() - document.left() - len) / 2.0F + 20F,document.bottom() - 15, 0);

        // 6.写入页脚2的模板（就是页脚的Y页这俩字）添加到文档中，计算模板的和Y轴,X=(右边界-左边界 - 前半部分的len值)/2.0F +
        // len ， y 轴和之前的保持一致，底边界-25
        cb.addTemplate(template,(document.rightMargin() + document.right() + document.leftMargin() - document.left()) / 2.0F ,document.bottom() - 15);
    }

    /**
     * @description: 关闭文档时，替换模板，完成整个页眉页脚组件
     *
     * @author: Fang Kun
     * @param writer
     * @param document
     * @date: 2020/3/18 13:28
     * @return: void
     */
    @Override
    public void onCloseDocument(PdfWriter writer, Document document) {
        template.beginText();
        if(ObjectUtil.isEmpty(bf)){
            initFront();
        }
        template.setFontAndSize(bf,fontSize);
        String replace=null;
        if(writer.getPageNumber()==1){
            replace = " 1 页";
        }else{
            replace= (writer.getPageNumber()-1) + " 页";
        }
        template.showText(replace);
        template.endText();
        template.closePath();
    }

    public static final String OS_NAME_FLAG_WIN = "win:\\";

    /**
     * @description: 初始化字体
     *
     * @author: Fang Kun
     * @date: 2020/3/18 13:28
     * @return: void
     */
    private void initFront(){
        if(StringUtils.isEmpty(fontFileName)){
            throw new RuntimeException("文件名称不能为空！");
        }
        try {
            if (bf == null) {
                // 判断这个路径后缀有没有 linux:/ win:\\
                // 判断系统的类型,是windows 还是linux
//                String property = System.getProperty(ImsConstants.OS_NAME);
//                String path = pdfMulPageDTO.getPath();
//                if(property.indexOf(ImsConstants.OS_NAME_FLAG_WIN)>-1 && path.lastIndexOf(ImsConstants.WINDOWS_SPLIT)==-1 && path.lastIndexOf(ImsConstants.OBLIQUE_LINE)==-1){
//                    // 是win系统且后缀没有 \\
//                    path=path+ImsConstants.WINDOWS_SPLIT;
//                }else{
//                    if(path.lastIndexOf(ImsConstants.OBLIQUE_LINE)==-1){
//                        path=path+ImsConstants.OBLIQUE_LINE;
//                    }
//                }
                //添加字体，以支持中文
                String fontPath =pdfMulPageDTO.getPath()+fontFileName;
                LOGGER.debug("initFront: fontPath--> "+fontPath);
                //创建基础字体
                bf = BaseFont.createFont(fontPath, BaseFont.IDENTITY_H, BaseFont.NOT_EMBEDDED);
                LOGGER.debug("BaseFont.createFont--> "+ JSON.toJSONString(bf.getFontType()));
            }
            if (fontDetail == null) {
                //数据体字体
                fontDetail = new Font(bf, fontSize, Font.NORMAL);
                LOGGER.debug("fontDetail--> "+ JSON.toJSONString(fontDetail));
            }
        } catch (DocumentException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

}