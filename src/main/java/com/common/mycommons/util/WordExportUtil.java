package com.common.mycommons.util;

import cn.afterturn.easypoi.word.entity.MyXWPFDocument;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.usermodel.Range;
import org.apache.poi.wp.usermodel.HeaderFooterType;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import javax.validation.constraints.NotNull;
import java.io.InputStream;
import java.math.BigInteger;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

public class WordExportUtil {
    private static final Logger LOGGER = LoggerFactory.getLogger(WordExportUtil.class);
    /** 1厘米 */
    public static final int ONE_UNIT = 567;
    /** 页脚样式 */
    public static final String STYLE_FOOTER = "footer";

    /** 页眉样式 */
    public static final String STYLE_HEADER = "header";
    /** 语言，简体中文 */
    public static final String LANG_ZH_CN = "zh-CN";

    /**
     * 导出word
     * @param fileName 文件输出名字
     * @param response 响应
     */
    public static void wordExport(String fileName, HttpServletResponse response) {
        // 新建一个文档
        XWPFDocument doc = new XWPFDocument();
        // ========================= 页眉 ==================================
        // 创建一个标题
        XWPFParagraph paragraph = doc.createParagraph();
        addCustomHeadingStyle(doc, "标题 1", 1);
        // 段落设置样式
        paragraph.setStyle("标题 1");
        // 创建样式
        XWPFRun paragraphRun = paragraph.createRun();
        // 加粗
        paragraphRun.setBold(true);
        // 颜色
        paragraphRun.setColor("000000");
        // 字体
        paragraphRun.setFontFamily("宋体");
        // 大小
        paragraphRun.setFontSize(20);
        //设置行间距
        paragraphRun.setTextPosition(35);
        // 设置文本
        paragraphRun.setText("这是一个标题");
        // ========================= 页眉 ===================================
        // 创建页眉
        createDefaultHeader(doc, "这是页眉");
        // ========================= 段落 ===================================
        // 创建一个段落
        XWPFParagraph para = doc.createParagraph();
        // 一个XWPFRun代表具有相同属性的一个区域。就是说在word里面给这个地方加上特定的样式
        XWPFRun run = para.createRun();
        // 加粗
        run.setBold(true);
        // 设置内容
        run.setText("这是段落内容....");
        // 设置字体
        run.setFontFamily("宋体");
        // 首行缩进
        para.setIndentationFirstLine(ONE_UNIT);
        // ========================= 表格 ===================================
        // 创建一个5行5列的表格  这里可以自定义行 列
        XWPFTable table = doc.createTable(5, 5);
        // 这里增加的列原本初始化创建的那5行在通过getTableCells()方法获取时获取不到，但通过row新增的就可以。
        // 给表格增加一列，变成6列
        // table.addNewCol();
        // 给表格新增一行，变成6行
        // table.createRow();
        List<XWPFTableRow> rows = table.getRows();
        //表格属性
        CTTblPr tablePr = table.getCTTbl().addNewTblPr();
        //表格宽度
        CTTblWidth width = tablePr.addNewTblW();
        width.setW(BigInteger.valueOf(8000));
        XWPFTableRow row;
        List<XWPFTableCell> cells;
        XWPFTableCell cell;
        int rowSize = rows.size();
        int cellSize;
        for (int i = 0; i < rowSize; i++) {
            row = rows.get(i);
            //新增单元格
            row.addNewTableCell();
            //设置行的高度
            row.setHeight(500);
            //行属性
//            CTTrPr rowPr = row.getCtRow().addNewTrPr();
            //这种方式是可以获取到新增的cell的。
//            List<CTTc> list = row.getCtRow().getTcList();
            cells = row.getTableCells();
            cellSize = cells.size();
            for (int j = 0; j < cellSize; j++) {
                cell = cells.get(j);
                if ((i + j) % 2 == 0) {
                    //设置单元格的颜色
                    //红色
                    cell.setColor("ff0000");
                } else {
                    //蓝色
                    cell.setColor("0000ff");
                }
                //单元格属性
                CTTcPr cellPr = cell.getCTTc().addNewTcPr();
                cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
                if (j == 3) {
                    //设置宽度
                    cellPr.addNewTcW().setW(BigInteger.valueOf(3000));
                }
                // 设置文本
                cell.setText(i + ", " + j);
            }
        }
        // 创建第二个段落
        XWPFParagraph paraTwo = doc.createParagraph();
        // 一个XWPFRun代表具有相同属性的一个区域。就是说在word里面给这个地方加上特定的样式
        XWPFRun runTwo = paraTwo.createRun();
        // 加粗
        runTwo.setBold(true);
        // 设置内容
        runTwo.setText("这是一个标题");
        /**
         * 字号「八号」对应磅值5
         * 字号「七号」对应磅值5.5
         * 字号「小六」对应磅值6.5
         * 字号「六号」对应磅值7.5
         * 字号「小五」对应磅值9
         * 字号「五号」对应磅值10.5
         * 字号「小四」对应磅值12
         * 字号「四号」对应磅值14
         * 字号「小三」对应磅值15
         * 字号「三号」对应磅值16
         * 字号「小二」对应磅值18
         * 字号「二号」对应磅值22
         * 字号「小一」对应磅值24
         * 字号「一号」对应磅值26
         * 字号「小初」对应磅值36
         * 字号「初号」对应磅值42
         */
        runTwo.setFontSize(22);

        // 创建页脚
        createDefaultFooter(doc);

        // 将doc 写入输出流中
        ServletOutputStream outputStream =null;
        try {
            PdfMulPageExportUtils.responseHandler(response, fileName);
            outputStream = response.getOutputStream();
            doc.write(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }

        // 关流
        FileUtiles.closeStreamOfOut(outputStream);
    }

    /**
     * 增加自定义标题样式。
     *
     * @param docxDocument 目标文档
     * @param strStyleId 样式名称
     * @param headingLevel 样式级别
     */
    private static void addCustomHeadingStyle(@NotNull XWPFDocument docxDocument, @NotNull String strStyleId, @NotNull int headingLevel) {

        CTStyle ctStyle = CTStyle.Factory.newInstance();
        ctStyle.setStyleId(strStyleId);

        CTString styleName = CTString.Factory.newInstance();
        styleName.setVal(strStyleId);
        ctStyle.setName(styleName);

        CTDecimalNumber indentNumber = CTDecimalNumber.Factory.newInstance();
        indentNumber.setVal(BigInteger.valueOf(headingLevel));

        // lower number > style is more prominent in the formats bar
        ctStyle.setUiPriority(indentNumber);

        CTOnOff onoffnull = CTOnOff.Factory.newInstance();
        ctStyle.setUnhideWhenUsed(onoffnull);

        // style shows up in the formats bar
        ctStyle.setQFormat(onoffnull);

        // style defines a heading of the given level
        CTPPr ppr = CTPPr.Factory.newInstance();
        ppr.setOutlineLvl(indentNumber);
        ctStyle.setPPr(ppr);

        XWPFStyle style = new XWPFStyle(ctStyle);

        // is a null op if already defined
        XWPFStyles styles = docxDocument.createStyles();
        style.setType(STStyleType.PARAGRAPH);
        styles.addStyle(style);
    }


    /**
     * 创建默认页眉
     *
     * @param docx XWPFDocument文档对象
     * @param text 页眉文本
     * @return 返回文档帮助类对象，可用于方法链调用
     */
    private static void createDefaultHeader(final XWPFDocument docx, final String text){
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph paragraph = new XWPFParagraph(ctp, docx);
        ctp.addNewR().addNewT().setStringValue(text);
        ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
        XWPFHeader header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[] { paragraph });
        header.setXWPFDocument(docx);
    }

    /**
     * 创建默认的页脚(该页脚主要只居中显示页码)
     *
     * @param docx
     *            XWPFDocument文档对象
     * @return 返回文档帮助类对象，可用于方法链调用
     */
    public static void createDefaultFooter(final XWPFDocument docx) {
        // TODO 设置页码起始值
        CTP pageNo = CTP.Factory.newInstance();
        XWPFParagraph footer = new XWPFParagraph(pageNo, docx);
        CTPPr begin = pageNo.addNewPPr();
        begin.addNewPStyle().setVal(STYLE_FOOTER);
        begin.addNewJc().setVal(STJc.CENTER);
        pageNo.addNewR().addNewFldChar().setFldCharType(STFldCharType.BEGIN);
        pageNo.addNewR().addNewInstrText().setStringValue("PAGE   \\* MERGEFORMAT");
        pageNo.addNewR().addNewFldChar().setFldCharType(STFldCharType.SEPARATE);
        CTR end = pageNo.addNewR();
        CTRPr endRPr = end.addNewRPr();
        endRPr.addNewNoProof();
        endRPr.addNewLang().setVal(LANG_ZH_CN);
        end.addNewFldChar().setFldCharType(STFldCharType.END);
        CTSectPr sectPr = docx.getDocument().getBody().isSetSectPr() ? docx.getDocument().getBody().getSectPr() : docx.getDocument().getBody().addNewSectPr();
        XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(docx, sectPr);
        policy.createFooter(STHdrFtr.DEFAULT, new XWPFParagraph[] { footer });
    }

    /***
     *
     * @param outFileName 文件输出名字
     * @param fileName 模板文件名字
     * @param response 响应
     */
    public static void wordExportForTemplates(String outFileName, String fileName, Map<String, Object> data, HttpServletResponse response){
        Objects.requireNonNull(fileName, "file path is not must be null");
        Objects.requireNonNull(outFileName, "file name is not must be null");
        Objects.requireNonNull(data, "file data is not must be null");
        InputStream resourceAsStream = Thread.currentThread().getContextClassLoader().getResourceAsStream(fileName);
        ServletOutputStream outputStream = null;
        try {
            HWPFDocument doc = new HWPFDocument(resourceAsStream);
            Range range = doc.getRange();
            /**
             * 模板中的每个占位符都是 ${key}
             * 遍历map就可以了
             */
            Set<Map.Entry<String, Object>> entries = data.entrySet();
            for (Map.Entry<String, Object> entry : entries) {
                range.replaceText("${" + entry.getKey() + "}", entry.getValue().toString());
            }
            outputStream = response.getOutputStream();
            //把doc输出到输出流中
            doc.write(outputStream);
        }catch (Exception e){
            e.printStackTrace();
        }
        FileUtiles.closeStreamOfOut(outputStream);
    }

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
            cn.afterturn.easypoi.word.WordExportUtil.exportWord07(myXWPFDocument, data);
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
