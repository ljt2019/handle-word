package com.tiger.poi.xwpf;

import org.apache.commons.io.IOUtils;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.Document;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
import org.apache.xmlbeans.XmlOptions;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTBody;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 合并两个docx文档方法，支持图片合并,【poi-ooxml-3.13.jar】
 * Create by tiger on 2019/5/5
 */
public class POIMergeDocUtil {

    public static void main(String[] args) throws Exception {
        String src1 = "C:\\Users\\tiger\\Desktop\\test1picture.docx";
        String src2 = "C:\\Users\\tiger\\Desktop\\test2pitucre.docx";
        String src3 = "C:\\Users\\tiger\\Desktop\\test3.docx";
        String src4 = "C:\\Users\\tiger\\Desktop\\test4picture.doc";
        String src5 = "C:\\Users\\tiger\\Desktop\\test5.doc";
        String[] srcDocxs =
                new String[]{src1, src4, src5};
        String dest = "C:\\Users\\tiger\\Desktop\\dest.doc";
//        mergeDoc(srcDocxs, dest);
        mergeWord(srcDocxs, dest);

    }

    /**
     * 支持图片合并，源文件支持 .docx,输出文件兼容doc和docx
     *
     * @param srcDocxs
     * @param destDocx
     */
    public static void mergeWord(String[] srcDocxs, String destDocx) {
        OutputStream dest = null;
        List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
        int length = null == srcDocxs ? 0 : srcDocxs.length;
        /**
         * 循环获取每个docx文件的OPCPackage对象
         */
        for (int i = 0; i < length; i++) {
            String doc = srcDocxs[i];
            OPCPackage srcPackage = null;
            try {
                srcPackage = OPCPackage.open(doc);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (null != srcPackage) {
                opcpList.add(srcPackage);
            }
        }
        int opcpSize = opcpList.size();
        //获取的OPCPackage对象大于0时，执行合并操作
        if (opcpSize > 0) {
            try {
                dest = new FileOutputStream(destDocx);
                XWPFDocument srcDoc = new XWPFDocument(opcpList.get(0));
                //OPCPackage大于1的部分执行合并操作
                if (opcpSize > 1) {
                    for (int i = 1; i < opcpSize; i++) {
                        OPCPackage src2Package = opcpList.get(i);
                        XWPFDocument appendDoc =
                                new XWPFDocument(src2Package);
                        appendDoc(srcDoc, appendDoc);
                    }
                }
                //将合并的文档写入目标文件中
                srcDoc.write(dest);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                IOUtils.closeQuietly(dest);
            }
        }
    }

    /**
     * 不支持图片，源文件支持 .docx,输出文件兼容doc和docx
     *
     * @param srcDocxs 需要合并的目标docx文件
     * @param destDocx 合并后的docx输出文件
     */
    public static void mergeDoc(String[] srcDocxs, String destDocx) {

        OutputStream dest = null;
        List<OPCPackage> opcpList = new ArrayList<OPCPackage>();
        int length = null == srcDocxs ? 0 : srcDocxs.length;
        /**
         * 循环获取每个docx文件的OPCPackage对象
         */
        for (int i = 0; i < length; i++) {
            String doc = srcDocxs[i];
            OPCPackage srcPackage = null;
            try {
                srcPackage = OPCPackage.open(doc);
            } catch (Exception e) {
                e.printStackTrace();
            }
            if (null != srcPackage) {
                opcpList.add(srcPackage);
            }
        }

        int opcpSize = opcpList.size();
        //获取的OPCPackage对象大于0时，执行合并操作
        if (opcpSize > 0) {
            try {
                dest = new FileOutputStream(destDocx);
                XWPFDocument src1Document = new XWPFDocument(opcpList.get(0));
                CTBody src1Body = src1Document.getDocument().getBody();
                //OPCPackage大于1的部分执行合并操作
                if (opcpSize > 1) {
                    for (int i = 1; i < opcpSize; i++) {
                        OPCPackage src2Package = opcpList.get(i);
                        XWPFDocument src2Document =
                                new XWPFDocument(src2Package);
                        CTBody src2Body = src2Document.getDocument().getBody();
                        appendBody(src1Body, src2Body);
                    }
                }
                //将合并的文档写入目标文件中
                src1Document.write(dest);
            } catch (FileNotFoundException e) {
                e.printStackTrace();
            } catch (IOException e) {
                e.printStackTrace();
            } catch (Exception e) {
                e.printStackTrace();
            } finally {
                IOUtils.closeQuietly(dest);
            }
        }
    }

    /**
     * 支持图片
     * 两个XWPFDocument对象进行追加，含有图片
     * XWPFDocument xmd = null;
     * for (int i = 0; i < list.size(); i++) {
     * xmd = list.get(0);
     * if (i != 0) {
     * POIMergeDocUtil.appendBody(xmd, list.get(i));
     * }
     * }
     *
     * @param srcDoc    源文件
     * @param appendDoc 追加文件
     * @throws Exception
     */
    public static void appendDoc(XWPFDocument srcDoc,
                                 XWPFDocument appendDoc) throws Exception {
        CTBody srcBody = srcDoc.getDocument().getBody();
        CTBody appendBody = appendDoc.getDocument().getBody();

        List<XWPFPictureData> allPictures = appendDoc.getAllPictures();
        // 记录图片合并前及合并后的ID
        Map<String, String> map = new HashMap();
        for (XWPFPictureData picture : allPictures) {
            String before = appendDoc.getRelationId(picture);
            //将原文档中的图片加入到目标文档中
            String after =
                    srcDoc.addPictureData(picture.getData(), Document.PICTURE_TYPE_PNG);
            map.put(before, after);
        }
        appendBody(srcBody, appendBody, map);
    }

    /**
     * 支持图片
     *
     * @param src
     * @param append
     * @param map
     * @throws Exception
     */
    private static void appendBody(CTBody src, CTBody append,
                                   Map<String, String> map) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);

        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart =
                srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart =
                appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        if (map != null && !map.isEmpty()) {
            //对xml字符串中图片ID进行替换
            for (Map.Entry<String, String> set : map.entrySet()) {
                addPart = addPart.replace(set.getKey(), set.getValue());
            }
        }
        //将两个文档的xml内容进行拼接
        CTBody makeBody =
                CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src.set(makeBody);
    }


    /**
     * 不支持图片
     * XWPFDocument xmd = list.get(0); //默认获取第一个作为模板
     * for (int i = 0; i < list.size(); i++) {
     * if (i == 0) {
     * xmd = list.get(0);
     * continue;
     * } else {
     * //当存在多条时候
     * xmd = POIMergeDocUtil.mergeWord(xmd, list.get(i));
     * }
     * }
     *
     * @param document
     * @param doucDocument2
     * @return
     * @throws Exception
     */
    public static XWPFDocument mergeDoc(XWPFDocument document,
                                        XWPFDocument doucDocument2) throws Exception {
        XWPFDocument src1Document = document;
        //        XWPFParagraph p = src1Document.createParagraph();
        //        p.setPageBreak(true);//设置分页符
        CTBody src1Body = src1Document.getDocument().getBody();
        XWPFDocument src2Document = doucDocument2;
        CTBody src2Body = src2Document.getDocument().getBody();
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = src2Body.xmlText(optionsOuter);
        String srcString = src1Body.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart =
                srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart =
                appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody =
                CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src1Body.set(makeBody);
        return src1Document;
    }

    /**
     * 不支持图片
     *
     * @param src    目标文档
     * @param append 要合并的文档
     * @throws Exception
     */
    private static void appendBody(CTBody src,
                                   CTBody append) throws Exception {
        XmlOptions optionsOuter = new XmlOptions();
        optionsOuter.setSaveOuter();
        String appendString = append.xmlText(optionsOuter);
        String srcString = src.xmlText();
        String prefix = srcString.substring(0, srcString.indexOf(">") + 1);
        String mainPart =
                srcString.substring(srcString.indexOf(">") + 1, srcString.lastIndexOf("<"));
        String sufix = srcString.substring(srcString.lastIndexOf("<"));
        String addPart =
                appendString.substring(appendString.indexOf(">") + 1, appendString.lastIndexOf("<"));
        CTBody makeBody =
                CTBody.Factory.parse(prefix + mainPart + addPart + sufix);
        src.set(makeBody);
    }
}