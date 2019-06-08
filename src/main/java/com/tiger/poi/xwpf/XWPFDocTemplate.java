package com.tiger.poi.xwpf;

import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTFonts;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTP;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.STMerge;

import java.io.*;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 仅支持对docx文件的文本及表格中的内容进行替换
 * Create by tiger on 2019/5/5
 */
public class XWPFDocTemplate {

    public static void main(String[] args) {
        String src = "C:\\Users\\tiger\\Desktop\\联奕科技\\test3.docx";
        String dest = "C:\\Users\\tiger\\Desktop\\联奕科技\\dest3.doc";
        InputStream is = null;
        OutputStream os = null;
        try {
            is = new FileInputStream(src);
            os = new FileOutputStream(dest);
            XWPFDocument doc = new XWPFDocument(is); //文档输入流
            Map<String, String> map = new HashMap<>();
            map.put("name", "李工");
            map.put("age", "25");
            map.put("title", "标题侧部");
            map.put("major", "主修专业");
            replaceParagraphs(doc, map);
            replaceTable(doc, 0, map);
            doc.write(os);

            os.flush();
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        } finally {
            if (os != null) {
                try {
                    os.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
            if (is != null) {
                try {
                    is.close();
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
    }


    /**
     * 替换指定表格中的标签，key【匹配模板中的${key}】
     *
     * @param map key(待替换标签)-value(文本内容)
     */
    public static void replaceTable(XWPFDocument document, int tableIndex,
                                    Map<String, String> map) {
        List<XWPFTable> tables = document.getTables();
        XWPFTable table = tables.get(tableIndex);
        new XWPFTableHandler(table).replaceParagraphs(map);
    }

    /**
     * 替换所有表格中的标签，key【匹配模板中的${key}】
     *
     * @param map key(待替换标签)-value(文本内容)
     */
    public static void replaceAllTable(XWPFDocument document,
                                       Map<String, String> map) {
        List<XWPFTable> tables = document.getTables();
        for (XWPFTable table : tables) {
            new XWPFTableHandler(table).replaceParagraphs(map);
        }
    }

    /**
     * 替换文本中的标签
     *
     * @param map key(待替换标签)-value(文本内容)
     */
    public static void replaceParagraphs(XWPFDocument document,
                                         Map<String, String> map) {
        List<XWPFParagraph> paragraphs = document.getParagraphs();
        for (XWPFParagraph paragraph : paragraphs) {
            new XWPFParagraphHandler(paragraph).replaceAll(map);
        }
    }

    /**
     * 为表格中的单元格赋值
     *
     * @param cell
     * @param value     替换的值
     * @param paraIndex 段落索引
     * @param runIndex
     * @param fontSize  字体大小
     * @param typeface  字体类型，默认宋体正文
     */
    public static void setCellText(XWPFTableCell cell,
                                   String value, int fontSize, String typeface,
                                   int paraIndex, int runIndex) {
        List<XWPFParagraph> paras = cell.getParagraphs(); //获取行中某个表格中所有段落
        XWPFParagraph para = paras.get(paraIndex); //获取某段落
        List<XWPFRun> runs = para.getRuns();
        XWPFRun run;
        if (runs.size() <= 0) {
            run = para.insertNewRun(runIndex);
        } else {
            run = runs.get(runIndex);
        }
        if (typeface == null) {
            typeface = "宋体正文";
        }
        run.setFontFamily(typeface);
        run.setFontSize(fontSize);
        run.setText(value);
    }

    /**
     * 合并单元格(水平合并)
     *
     * @param table
     * @param row      行索引
     * @param fromCell 起始列
     * @param toCell   结束列
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row,
                                            int fromCell, int toCell) {
        for (int cellIndex = fromCell; cellIndex <= toCell; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCell) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewHMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    /**
     * 合并单元格(纵向合并)
     *
     * @param table
     * @param col     列索引
     * @param fromRow 起始行
     * @param toRow   结束行
     */
    public static void mergeCellsVertically(XWPFTable table, int col,
                                            int fromRow, int toRow) {
        for (int rowIndex = fromRow; rowIndex <= toRow; rowIndex++) {
            XWPFTableCell cell = table.getRow(rowIndex).getCell(col);
            if (rowIndex == fromRow) {
                // The first merged cell is set with RESTART merge value
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.RESTART);
            } else {
                // Cells which join (merge) the first one, are set with CONTINUE
                cell.getCTTc().addNewTcPr().addNewVMerge().setVal(STMerge.CONTINUE);
            }
        }
    }

    public static void getParagraph(XWPFTableCell cell, String cellText) {
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph p = new XWPFParagraph(ctp, cell);
        p.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = p.createRun();
        run.setText(cellText);
        CTRPr rpr =
                run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
        CTFonts fonts =
                rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
        fonts.setAscii("仿宋");
        fonts.setEastAsia("仿宋");
        fonts.setHAnsi("仿宋");
        cell.setParagraph(p);
    }

    /**
     * 在doc中插入图片
     *
     * @param picSuffix 图片类型,图片后缀，例如【jpg】
     * @param width     宽
     * @param height    高
     * @param paragraph 段落，图片插入的位置
     */
    public static void createPicture(XWPFDocument doc, XWPFParagraph paragraph,
                                     String picSuffix, int width, int height) {
        int pictureType = getPictureType(picSuffix);
        final int EMU = 9525;
        width *= EMU;
        height *= EMU;
        String blipId =
                doc.getAllPictures().get(pictureType).getPackageRelationship().getId();
        CTInline inline =
                paragraph.createRun().getCTR().addNewDrawing().addNewInline();
        String picXml =
                "" + "<a:graphic xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\">" +
                        "   <a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "      <pic:pic xmlns:pic=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
                        "         <pic:nvPicPr>" + "            <pic:cNvPr id=\"" + pictureType +
                        "\" name=\"Generated\"/>" + "            <pic:cNvPicPr/>" +
                        "         </pic:nvPicPr>" + "         <pic:blipFill>" +
                        "            <a:blip r:embed=\"" + blipId +
                        "\" xmlns:r=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships\"/>" +
                        "            <a:stretch>" + "               <a:fillRect/>" +
                        "            </a:stretch>" + "         </pic:blipFill>" +
                        "         <pic:spPr>" + "            <a:xfrm>" +
                        "               <a:off x=\"0\" y=\"0\"/>" +
                        "               <a:ext cx=\"" + width + "\" cy=\"" + height +
                        "\"/>" + "            </a:xfrm>" +
                        "            <a:prstGeom prst=\"rect\">" +
                        "               <a:avLst/>" + "            </a:prstGeom>" +
                        "         </pic:spPr>" + "      </pic:pic>" +
                        "   </a:graphicData>" + "</a:graphic>";

        inline.addNewGraphic().addNewGraphicData();
        XmlToken xmlToken = null;
        try {
            xmlToken = XmlToken.Factory.parse(picXml);
        } catch (XmlException xe) {
            xe.printStackTrace();
        }
        inline.set(xmlToken);
        inline.setDistT(0);
        inline.setDistB(0);
        inline.setDistL(0);
        inline.setDistR(0);
        CTPositiveSize2D extent = inline.addNewExtent();
        extent.setCx(width);
        extent.setCy(height);
        CTNonVisualDrawingProps docPr = inline.addNewDocPr();
        docPr.setId(pictureType);
        docPr.setName("图片名称" + pictureType);
        docPr.setDescr("图片描述");
    }

    /**
     * 获取图片类型
     *
     * @param picType
     * @return
     */
    private static int getPictureType(String picType) {
        int res = XWPFDocument.PICTURE_TYPE_PICT;
        if (picType != null) {
            if (picType.equalsIgnoreCase("png")) {
                res = XWPFDocument.PICTURE_TYPE_PNG;
            } else if (picType.equalsIgnoreCase("dib")) {
                res = XWPFDocument.PICTURE_TYPE_DIB;
            } else if (picType.equalsIgnoreCase("emf")) {
                res = XWPFDocument.PICTURE_TYPE_EMF;
            } else if (picType.equalsIgnoreCase("jpg") ||
                    picType.equalsIgnoreCase("jpeg")) {
                res = XWPFDocument.PICTURE_TYPE_JPEG;
            } else if (picType.equalsIgnoreCase("wmf")) {
                res = XWPFDocument.PICTURE_TYPE_WMF;
            }
        }
        return res;
    }

}
