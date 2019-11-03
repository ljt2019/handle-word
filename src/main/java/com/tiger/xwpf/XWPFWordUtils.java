package com.tiger.xwpf;

import com.tiger.enums.XWPFBorderStyleEnum;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.xwpf.model.XWPFHeaderFooterPolicy;
import org.apache.poi.xwpf.usermodel.*;
import org.apache.xmlbeans.XmlException;
import org.apache.xmlbeans.XmlToken;
import org.apache.xmlbeans.impl.xb.xmlschema.SpaceAttribute;
import org.openxmlformats.schemas.drawingml.x2006.main.CTNonVisualDrawingProps;
import org.openxmlformats.schemas.drawingml.x2006.main.CTPositiveSize2D;
import org.openxmlformats.schemas.drawingml.x2006.wordprocessingDrawing.CTInline;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.*;

import java.io.IOException;
import java.math.BigInteger;
import java.util.List;

/**
 * 仅支持对docx文件的文本及表格中的内容进行替换
 * 以及向表格中指定位置插入内容（x,y）
 */
public class XWPFWordUtils {

    private XWPFWordUtils() {
    }

    /**
     * 设置页面大小及纸张方向 landscape横向
     * 信纸:("15840", "12240")
     * A4:("16837", "11905")
     *
     * @param doc
     * @param width
     * @param height
     * @param stValue
     */
    public static void setDocSize(XWPFDocument doc, String width, String height, STPageOrientation.Enum stValue) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageSz pgsz = sectPr.isSetPgSz() ? sectPr.getPgSz() : sectPr.addNewPgSz();
        pgsz.setH(new BigInteger(height));
        pgsz.setW(new BigInteger(width));
        pgsz.setOrient(stValue);
    }

    /**
     * 设置页边距 (word中1厘米约等于567)
     *
     * @param doc
     * @param left
     * @param top
     * @param right
     * @param bottom
     */
    public static void setDocMargin(XWPFDocument doc, String left, String top, String right, String bottom) {
        CTSectPr sectPr = doc.getDocument().getBody().addNewSectPr();
        CTPageMar ctpagemar = sectPr.addNewPgMar();
        if (StringUtils.isNotBlank(left)) {
            ctpagemar.setLeft(new BigInteger(left));
        }
        if (StringUtils.isNotBlank(top)) {
            ctpagemar.setTop(new BigInteger(top));
        }
        if (StringUtils.isNotBlank(right)) {
            ctpagemar.setRight(new BigInteger(right));
        }
        if (StringUtils.isNotBlank(bottom)) {
            ctpagemar.setBottom(new BigInteger(bottom));
        }
    }

    /**
     * 创建默认页眉
     *
     * @param doc  档对象
     * @param text 页眉文本
     * @throws IOException
     */
    public static void createDefaultHeader(final XWPFDocument doc, final String text) {
        try {
            CTP ctp = CTP.Factory.newInstance();
            XWPFParagraph paragraph = new XWPFParagraph(ctp, doc);
            paragraph.setAlignment(ParagraphAlignment.CENTER);
            ctp.addNewR().addNewT().setStringValue(text);
            ctp.addNewR().addNewT().setSpace(SpaceAttribute.Space.PRESERVE);
            CTSectPr sectPr = doc.getDocument().getBody().isSetSectPr() ? doc.getDocument().getBody().getSectPr() : doc.getDocument().getBody().addNewSectPr();
            XWPFHeaderFooterPolicy policy = new XWPFHeaderFooterPolicy(doc, sectPr);
            XWPFHeader header = null;
            header = policy.createHeader(STHdrFtr.DEFAULT, new XWPFParagraph[]{paragraph});
            header.setXWPFDocument(doc);
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * 分页
     *
     * @param doc       XWPFDocument
     * @param breakType 分页类型
     */
    public static void addNewPage(XWPFDocument doc, BreakType breakType) {
        XWPFParagraph xp = doc.createParagraph();
        xp.createRun().addBreak(breakType);
    }

    /**
     * 获得一个指定行数和列数的表格
     *
     * @param doc  XWPFDocument 对象
     * @param rows 行数
     * @param cols 列数
     * @return 表格对象
     */
    public static XWPFTable getTable(XWPFDocument doc, int rows, int cols, boolean style) {
        XWPFTable table = doc.createTable(rows, cols);
        if (style) {
            setTableStyle(table, true, true);
        }
        table.setWidth(100);
        return table;
    }

    /**
     * 设置表格样式，表格线的样式与线宽大小
     *
     * @param table   XWPFTable 对象
     * @param inside  内线
     * @param outside 外线
     */
    public static void setTableStyle(XWPFTable table, boolean inside, boolean outside) {
        //线宽实体
        CTTblBorders borders = table.getCTTbl().getTblPr().addNewTblBorders();
        if (inside) {
            //表格内部横向线宽
            CTBorder hBorder = borders.addNewInsideH();
            hBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            hBorder.setSz(new BigInteger("1"));
//        hBorder.setColor("0000FF");
//        hBorder.setColor("00FF00");

            //表格内部纵向线宽
            CTBorder vBorder = borders.addNewInsideV();
            vBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            vBorder.setSz(new BigInteger("1"));
//        vBorder.setColor("0000FF");

        }

        if (outside) {
            //表格外边框,左
            CTBorder lBorder = borders.addNewLeft();
            lBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            lBorder.setSz(new BigInteger("1"));

            //表格外边框,上
            CTBorder tBorder = borders.addNewTop();
            tBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            tBorder.setSz(new BigInteger("1"));

            //表格外边框,右
            CTBorder rBorder = borders.addNewRight();
            rBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            rBorder.setSz(new BigInteger("1"));

            //表格外边框,下
            CTBorder bBorder = borders.addNewBottom();
            bBorder.setVal(STBorder.Enum.forString(XWPFBorderStyleEnum.CENTER.getMsg()));
            bBorder.setSz(new BigInteger("1"));
        }
    }

    /**
     * 1、向表格中格子填充内容
     * 2、设置列宽
     * 3、默认水平居左
     *
     * @param cell    XWPFTableCell 对象
     * @param width   格子宽度，例如【"2400"】
     * @param content 文本
     * @param size    字体大小
     */
    public static XWPFTableCell fillCellLeft(XWPFTableCell cell, String width, String content, int size, boolean bold) {
        //设置表格中格子列宽
        setTableWith(cell, width);
        //加粗居左
        fillCell(cell, content, size, bold, ParagraphAlignment.LEFT);
        return cell;
    }

    /**
     * 填充表格头部
     * 设置列宽,默认水平居中
     *
     * @param cell    XWPFTableCell 对象
     * @param width   格子宽度
     * @param content 文本内容
     * @param size    字体大小
     */
    public static XWPFTableCell fillTableTitle(XWPFTableCell cell, String width, String content, int size, boolean bold) {
        //设置表格中格子列宽
        setTableWith(cell, width);
        //加粗居中
        fillCell(cell, content, size, bold, ParagraphAlignment.CENTER);
        return cell;
    }

    /**
     * 设置表格中格子列宽
     *
     * @param cell  格子
     * @param width 格子宽度 例如 【"2400"】
     */
    private static void setTableWith(XWPFTableCell cell, String width) {
        CTTc cttc = cell.getCTTc();
        CTTcPr cellPr = cttc.addNewTcPr();
        cellPr.addNewVAlign().setVal(STVerticalJc.CENTER);
        cttc.getPList().get(0).addNewPPr().addNewJc().setVal(STJc.CENTER);
        CTTblWidth tblWidth = cellPr.isSetTcW() ? cellPr.getTcW() : cellPr.addNewTcW();
        if (!StringUtils.isEmpty(width)) {
            tblWidth.setW(new BigInteger(width));
            tblWidth.setType(STTblWidth.DXA);
        }
    }

    /**
     * 向doc中填充文本
     *
     * @param doc     XWPFDocument 对象
     * @param content 文本内容
     * @param size    字体大小
     * @param bold    是否加粗
     * @return XWPFRun 对象
     */
    public static XWPFRun fillDoc(XWPFDocument doc, String content, int size, boolean bold) {
        XWPFParagraph xp = doc.createParagraph();
        XWPFRun run = xp.createRun();
        //设置行间距
        run.setTextPosition(1);
        //对齐方式
        xp.setAlignment(ParagraphAlignment.CENTER);
        //设置颜色--十六进制
        run.setColor("000000");
        //字体
        run.setFontFamily("宋体");
        //字体大小
        run.setFontSize(size);
        //加粗
        run.setBold(bold);
        //文本内容
        run.setText(content);
        return run;
    }

    /**
     * 向表格中的指定格子填充文本
     *
     * @param table   XWPFTable 对象
     * @param row     行索引
     * @param col     列索引
     * @param content 内容
     * @param size    字体大小
     */
    public static XWPFTableCell fillCell(XWPFTable table, int row, int col, String content, int size) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        //方式1
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        fillCell(cell, content, size);
        //方式2
//        fillCell(cell, content, size,false);
        return cell;
    }

    /**
     * 向表格中的指定格子填充文本
     *
     * @param table   XWPFTable 对象
     * @param row     行索引
     * @param col     列索引
     * @param content 文本内容
     * @param size    字体大小
     */
    public static XWPFTableCell fillCell(XWPFTable table, int row, int col, String content, int size, boolean bold) {
        XWPFTableCell cell = table.getRow(row).getCell(col);
        //方式1
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        fillCell(cell, content, size, bold, ParagraphAlignment.CENTER);
        //方式2
//        fillCell(cell, content, size,bold);
        return cell;
    }

    /**
     * 向表格中填充文本
     *
     * @param cell    XWPFTableCell对象
     * @param content 文本内容
     */
    public static XWPFTableCell fillCell(XWPFTableCell cell, String content, int size, boolean bold) {
        CTP ctp = CTP.Factory.newInstance();
        XWPFParagraph p = new XWPFParagraph(ctp, cell);
        p.setAlignment(ParagraphAlignment.CENTER);
        XWPFRun run = p.createRun();
        run.setText(content);
        run.setFontFamily("宋体");
        run.setFontSize(size);
        run.setBold(bold);
//        CTRPr rpr =  run.getCTR().isSetRPr() ? run.getCTR().getRPr() : run.getCTR().addNewRPr();
//        CTFonts fonts =  rpr.isSetRFonts() ? rpr.getRFonts() : rpr.addNewRFonts();
//        fonts.setAscii("仿宋");
//        fonts.setEastAsia("仿宋");
//        fonts.setHAnsi("仿宋");
        cell.setParagraph(p);
        cell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
        return cell;
    }

    /**
     * 合并单元格(水平合并)
     *
     * @param table   XWPFTable对象
     * @param row     行索引
     * @param fromCol 合并起始列索引
     * @param toCol   合并结束列索引
     */
    public static void mergeCellsHorizontal(XWPFTable table, int row, int fromCol, int toCol) {
        for (int cellIndex = fromCol; cellIndex <= toCol; cellIndex++) {
            XWPFTableCell cell = table.getRow(row).getCell(cellIndex);
            if (cellIndex == fromCol) {
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
     * @param table   XWPFTable 对象
     * @param col     列索引
     * @param fromRow 合并起始行索引
     * @param toRow   合并结束行索引
     */
    public static void mergeCellsVertically(XWPFTable table, int col, int fromRow, int toRow) {
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

    /**
     * 向格子中填充文本
     *
     * @param cell    XWPFTableCell 对象
     * @param content 文本内容
     * @param size    字体大小
     */
    private static void fillCell(XWPFTableCell cell, String content, int size) {
        //获取某段落
        XWPFParagraph para = cell.getParagraphs().get(0);
        //居中
        para.setAlignment(ParagraphAlignment.CENTER);
        fillRun(para, content, size);
    }

    /**
     * 向格子中填充文本
     *
     * @param cell      XWPFTableCell 对象
     * @param content   文本内容
     * @param size      字体大小
     * @param bold      是否加粗
     * @param alignment 例如 ParagraphAlignment.CENTER
     */
    private static void fillCell(XWPFTableCell cell, String content, int size, boolean bold, ParagraphAlignment alignment) {
        XWPFParagraph para = cell.getParagraphs().get(0);
        para.setAlignment(alignment);
        XWPFRun run = fillRun(para, content, size);
        run.setBold(bold);
    }

    /**
     * 向段落中填充文本
     *
     * @param para    XWPFParagraph 对象
     * @param content 文本内容
     * @param size    字体大小
     * @return XWPFRun 对象
     */
    private static XWPFRun fillRun(XWPFParagraph para, String content, int size) {
        XWPFRun run;
        List<XWPFRun> runs = para.getRuns();
        if (runs.size() <= 0) {
            run = para.insertNewRun(0);
            run.setText(content);
        } else {
            run = runs.get(0);
            run.setText(content, 0);
        }
        //宋体正文
        run.setFontFamily("宋体");
        run.setFontSize(size);
        //设置行间距
        run.setTextPosition(8);
        return run;
    }

    /*********************** 图片处理  *********************/

    /**
     * 获得图片类型
     *
     * @param picType
     * @return
     */
    public static int getPictureType(String picType) {
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

    /**
     * 向段落中插入图片
     *
     * @param pictureType
     * @param width       宽
     * @param height      高
     * @param paragraph   段落
     */
    public static void insertPicture(XWPFDocument doc, XWPFParagraph paragraph, int pictureType, int width, int height) {
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
        docPr.setName("图片类型" + pictureType);
        docPr.setDescr("图片描述");
    }
}
