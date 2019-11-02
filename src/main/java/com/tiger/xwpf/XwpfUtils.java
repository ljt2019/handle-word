package com.tiger.xwpf;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.util.Units;
import org.apache.poi.xwpf.usermodel.*;

import java.io.*;
import java.net.HttpURLConnection;
import java.net.URL;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

//import org.apache.poi.xwpf.usermodel.TableWidthType;

public class XwpfUtils {
    public static void main(String[] args) {
        // 文档生成方法
        XWPFDocument doc = new XWPFDocument();
        XWPFParagraph p1 = doc.createParagraph(); // 创建段落
        XWPFRun r1 = p1.createRun(); // 创建段落文本
        r1.setText("目录"); // 设置文本
        FileOutputStream out = null; // 创建输出流
        try {
            // 向word文档中添加内容
            XWPFParagraph p3 = doc.createParagraph(); // 创建段落
            XWPFRun r3 = p3.createRun(); // 创建段落文本
//            r3.addTab(); // tab
            r3.addBreak(); // 换行
            r3.setBold(true);
            XWPFParagraph p2 = doc.createParagraph(); // 创建段落
            XWPFRun r2 = p2.createRun(); // 创建段落文本
            // 设置文本
            r2.setText("表名");
            r2.setFontSize(14);
            r2.setBold(true);
            XWPFTable table1 = doc.createTable(8, 10);
//            table1.setWidthType(TableWidthType.AUTO);

            // 获取到刚刚插入的行
            XWPFTableRow row1 = table1.getRow(0);
            // 设置单元格内容
            row1.getCell(0).setText("字段名");
            row1.getCell(1).setText("字段说明");
            row1.getCell(2).setText("数据类型");
            row1.getCell(3).setText("长度");
            row1.getCell(4).setText("索引");
            row1.getCell(5).setText("是否为空");
            row1.getCell(6).setText("主键");
            row1.getCell(7).setText("外键");
            row1.getCell(8).setText("缺省值");
            row1.getCell(9).setText("备注");

            doc.setTable(0, table1);
            String filePath = "F:/worldTest/simple.docx";
            out = new FileOutputStream(new File(filePath));
            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {

            if (out != null) {
                try {
                    out.close();
                } catch (IOException e) {

                    e.printStackTrace();
                }
            }

        }
    }

    /**
     * 用一个docx文档作为模板，然后替换其中的内容，再写入目标文档中。
     *
     * @throws Exception
     */
    public void testTemplateWrite() throws Exception {
        Map<String, Object> params = new HashMap<String, Object>();
        params.put("reportDate", "2014-02-28");
        params.put("appleAmt", "100.00");
        params.put("bananaAmt", "200.00");
        params.put("totalAmt", "300.00");
        String filePath = "D:\\word\\template.docx";
        InputStream is = new FileInputStream(filePath);
        XWPFDocument doc = new XWPFDocument(is);
        //替换段落里面的变量
        this.replaceInPara(doc, params);
        //替换表格里面的变量
        this.replaceInTable(doc, params, 0);
        OutputStream os = new FileOutputStream("D:\\word\\write.docx");
        doc.write(os);
        this.close(os);
        this.close(is);
    }

    /**
     * 替换段落里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     */
    private void replaceInPara(XWPFDocument doc,
                               Map<String, Object> params) throws Exception {
        Iterator<XWPFParagraph> iterator = doc.getParagraphsIterator();
        XWPFParagraph para;
        while (iterator.hasNext()) {
            para = iterator.next();
            this.replaceInPara(para, params);
        }
    }

    /**
     * 替换段落里面的变量
     *
     * @param para   要替换的段落
     * @param params 参数
     * @throws Exception
     * @throws IOException
     * @throws InvalidFormatException
     */
    private boolean replaceInPara(XWPFParagraph para,
                                  Map<String, Object> params) throws Exception {
        boolean data = false;
        List<XWPFRun> runs;
        //有符合条件的占位符
        System.out.println("=====para.getParagraphText()======" +
                para.getParagraphText());
        if (this.matcher(para.getParagraphText()).find()) {
            runs = para.getRuns();
            data = true;
            Map<Integer, String> tempMap = new HashMap<Integer, String>();
            for (int i = 0; i < runs.size(); i++) {
                XWPFRun run = runs.get(i);
                String runText = run.toString();
                //以"$"开头
                boolean begin = runText.indexOf("$") > -1;
                boolean end = runText.indexOf("}") > -1;
                if (begin && end) {
                    tempMap.put(i, runText);
                    fillBlock(para, params, tempMap, i);
                    continue;
                } else if (begin && !end) {
                    tempMap.put(i, runText);
                    continue;
                } else if (!begin && end) {
                    tempMap.put(i, runText);
                    fillBlock(para, params, tempMap, i);
                    continue;
                } else {
                    if (tempMap.size() > 0) {
                        tempMap.put(i, runText);
                        continue;
                    }
                    continue;
                }
            }
        } else if (this.matcherRow(para.getParagraphText())) {
            runs = para.getRuns();
            data = true;
        }
        return data;
    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private boolean matcherRow(String str) {
        Pattern pattern =
                Pattern.compile("\\$\\[(.+?)\\]", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher.find();
    }

    /**
     * 填充run内容
     * @param para
     * @param params
     * @param tempMap
     * @param index
     * @throws InvalidFormatException
     * @throws IOException
     * @throws Exception
     */
    @SuppressWarnings("rawtypes")
    private void fillBlock(XWPFParagraph para, Map<String, Object> params,
                           Map<Integer, String> tempMap,
                           int index) throws InvalidFormatException,
            IOException, Exception {
        Matcher matcher;
        if (tempMap != null && tempMap.size() > 0) {
            String wholeText = "";
            List<Integer> tempIndexList = new ArrayList<Integer>();
            for (Map.Entry<Integer, String> entry : tempMap.entrySet()) {
                tempIndexList.add(entry.getKey());
                wholeText += entry.getValue();
            }
            if (wholeText.equals("")) {
                return;
            }
            matcher = this.matcher(wholeText);
            if (matcher.find()) {
                boolean isPic = false;
                int width = 0;
                int height = 0;
                int picType = 0;
                String path = null;
                String keyText =
                        matcher.group().substring(2, matcher.group().length() - 1);
                Object value = params.get(keyText);
                String newRunText = "";
                if (value instanceof String) {
                    newRunText = matcher.replaceFirst(String.valueOf(value));
                } else if (value instanceof Map) { //插入图片
                    isPic = true;
                    Map pic = (Map) value;
                    width = Integer.parseInt(pic.get("width").toString());
                    height = Integer.parseInt(pic.get("height").toString());
                    picType = getPictureType(pic.get("type").toString());
                    path = pic.get("path").toString();
                }

                //模板样式
                XWPFRun tempRun = null;
                // 直接调用XWPFRun的setText()方法设置文本时，在底层会重新创建一个XWPFRun，把文本附加在当前文本后面，
                // 所以我们不能直接设值，需要先删除当前run,然后再自己手动插入一个新的run。
                for (Integer pos : tempIndexList) {
                    tempRun = para.getRuns().get(pos);
                    tempRun.setText("", 0);
                }
                if (isPic) {
                    //addPicture方法的最后两个参数必须用Units.toEMU转化一下
                    //para.insertNewRun(index).addPicture(getPicStream(path), picType, "测试",Units.toEMU(width), Units.toEMU(height));
                    tempRun.addPicture(getPicStream(path), picType, "测试",
                            Units.toEMU(width),
                            Units.toEMU(height));
                } else {
                    //样式继承
                    if (newRunText.indexOf("\n") > -1) {
                        String[] textArr = newRunText.split("\n");
                        if (textArr.length > 0) {
                            //设置字体信息
                            String fontFamily = tempRun.getFontFamily();
                            int fontSize = tempRun.getFontSize();
                            //logger.info("------------------"+fontSize);
                            for (int i = 0; i < textArr.length; i++) {
                                if (i == 0) {
                                    tempRun.setText(textArr[0], 0);
                                } else {
                                    if (true) {
                                        XWPFRun newRun = para.createRun();
                                        //设置新的run的字体信息
                                        newRun.setFontFamily(fontFamily);
                                        if (fontSize == -1) {
                                            newRun.setFontSize(10);
                                        } else {
                                            newRun.setFontSize(fontSize);
                                        }
                                        newRun.addBreak();
                                        newRun.setText(textArr[i], 0);
                                    }
                                }
                            }
                        }
                    } else {
                        System.out.println("======newRunText=======" +
                                newRunText);
                        tempRun.setText(newRunText, 0);
                    }
                }
            }
            tempMap.clear();
        }
    }

    /**
     * 根据图片类型，取得对应的图片类型代码
     *
     * @param picType
     * @return int
     */
    private int getPictureType(String picType) {
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

    private InputStream getPicStream(String picPath) throws Exception {
        URL url = new URL(picPath);
        //打开链接
        HttpURLConnection conn = (HttpURLConnection) url.openConnection();
        //设置请求方式为"GET"
        conn.setRequestMethod("GET");
        //超时响应时间为5秒
        conn.setConnectTimeout(5 * 1000);
        //通过输入流获取图片数据
        InputStream is = conn.getInputStream();
        return is;
    }

    /**
     * 替换表格里面的变量
     *
     * @param doc    要替换的文档
     * @param params 参数
     * @throws Exception
     */
    private void replaceInTable(XWPFDocument doc, Map<String, Object> params,
                                int tableIndex) throws Exception {
        List<XWPFTable> tableList = doc.getTables();
        if (tableList.size() <= tableIndex) {
            throw new Exception("tableIndex对应的表格不存在");
        }
        XWPFTable table = tableList.get(tableIndex);
        List<XWPFTableRow> rows;
        List<XWPFTableCell> cells;
        List<XWPFParagraph> paras;
        rows = table.getRows();
        for (XWPFTableRow row : rows) {
            cells = row.getTableCells();
            for (XWPFTableCell cell : cells) {
                paras = cell.getParagraphs();
                for (XWPFParagraph para : paras) {
                    this.replaceInPara(para, params);
                    System.out.println("======params=====" + params);
                }
            }
        }
    }

    /**
     * 正则匹配字符串
     *
     * @param str
     * @return
     */
    private Matcher matcher(String str) {
        Pattern pattern =
                Pattern.compile("\\$\\{(.+?)\\}", Pattern.CASE_INSENSITIVE);
        Matcher matcher = pattern.matcher(str);
        return matcher;
    }

    /**
     * 关闭输入流
     *
     * @param is
     */
    private void close(InputStream is) {
        if (is != null) {
            try {
                is.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

    /**
     * 关闭输出流
     *
     * @param os
     */
    private void close(OutputStream os) {
        if (os != null) {
            try {
                os.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }
    }

}
