package com.word;

import com.tiger.xwpf.WordTagTemplate;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.TreeMap;

public class WordTagTemplateTest {

    public static void main(String[] args) throws Exception {
        test();
    }

    static int count1 = 17;
    static int count2 = 17;
    static int count3 = 17;
    static int count4 = 17;
    static int count5 = 7;
    static int count6 = 7;

    public static void test() throws Exception {
        String template = "doc/xjkxx_template.docx"; //模板路径
        String outPath = "doc/学籍卡信息tiger05.docx"; //写出路径
        //方式 1
        File file = new File(template);
        FileInputStream fis = new FileInputStream(file);
        XWPFDocument doc = new XWPFDocument(fis); //文档输入流

        //方式 2
        //        XWPFDocument doc =
        //            new XWPFDocument(POIXMLDocument.openPackage(template)); //文档输入流
        WordTagTemplate.replaceTag(doc, getStudentData()); //学生基本信息


        System.out.println("====== 导出成功 ======");
    }


    /**
     * 设置学生基本信息
     *
     * @return
     */
    public static Map<String, String> getStudentData() {
        Map<String, String> map = new HashMap();
        map.put("Xh", "130403021032");
        map.put("Yxmc", "计算机学院");
        map.put("Zyfxmc", "软件工程");
        map.put("Nj", "2016到");
        map.put("Bjmc", "2014助产1班");
        map.put("Pyccmc", "本科");
        map.put("Xb", "男");
        map.put("Xm", "佳佳");
        map.put("Cym", "");
        map.put("Jg", "广东广州");
        map.put("Mz", "汉");
        map.put("Sfzhm", "345456765645565456");
        map.put("Rxny", "20190909");
        map.put("Jtdz", "广东省湛江市建筑这hi额外hi跌回地诶嘿迪DWI");
        map.put("Kssj1", "2018-6-12");
        map.put("Kssj2", "2018-6-13");
        map.put("Kssj3", "2018-6-14");
        map.put("Kssj4", "2018-6-15");
        map.put("Kssj5", "2018-6-16");
        map.put("Jssj1", "2018-6-16");
        map.put("Jssj2", "2018-6-16");
        map.put("Jssj3", "2018-6-16");
        map.put("Jssj4", "2018-6-16");
        map.put("Jssj5", "2018-6-16");
        map.put("Xxxx1", "华南理工");
        map.put("Xxxx2", "北京师范");
        map.put("Xxxx3", "上海复旦问问单位单位单位单位的我吃完");
        map.put("Xxxx4", "2018-6-16");
        map.put("Xxxx5", "2018-6-16");
        map.put("Byxx", "华光女子高中");
        map.put("Rxqwhcd", "高中");
        map.put("Rxxs", "普通高考");
        map.put("Jtcy1", "呵呵");
        map.put("Jtcy2", "哈哈哈");
        map.put("Jtcy3", "佳佳");
        map.put("Jtcy4", "嘻嘻");
        map.put("Jkfs", "身体状况良好！");
        map.put("Sjts11", "2");
        map.put("Sjts12", "21");
        map.put("Sjts21", "21");
        map.put("Sjts22", "21");
        map.put("Sjts31", "21");
        map.put("Sjts32", "21");

        map.put("Bjts11", "21");
        map.put("Bjts12", "21");
        map.put("Bjts21", "21");
        map.put("Bjts22", "21");
        map.put("Bjts31", "21");
        map.put("Bjts32", "21");
        map.put("Kkts11", "21");
        map.put("Kkts12", "21");
        map.put("Kkts21", "21");
        map.put("Kkts22", "21");
        map.put("Kkts31", "21");
        map.put("Kkts32", "21");
        map.put("Cdts11", "21");
        map.put("Cdts12", "21");
        map.put("Cdts21", "21");
        map.put("Cdts22", "21");
        map.put("Cdts31", "21");
        map.put("Cdts32", "21");
        map.put("Ztts11", "21");
        map.put("Ztts12", "21");
        map.put("Ztts21", "21");
        map.put("Ztts22", "21");
        map.put("Ztts31", "21");
        map.put("Ztts32", "21");

        map.put("Xxyj", "准予毕业");
        map.put("Jfxq", "何时因何原因受过何种奖励我还好均为现金你家我家\n本无或处分");
        map.put("Jccfsj", "20190909929解除处分时间");
        map.put("Bysx", "毕业是hi是hi还是会im");
        map.put("Byyzsh", "43758683484653842748");
        map.put("Lxhqx", "去内存接收到了万事大吉实习实习工作");


        map.put("Xjydxq", "休学对接完IE申迪依次为四蒂此时端午节诶实地" + "你的为你多年未\n");

        map.put("Byjd",
                "新华社北京4月2日电  近日，中共中央总书记、国家主席、中央军委主席习近平对民政工作作出重要指示。习近平强调，近年来，民政系统认真贯彻中央决策部署，革弊鼎新、攻坚克难，各项事业取得新进展，有力服务了改革发展稳定大局。\n" +
                "\n" +
                "　　习近平指出，民政工作关系民生、连着民心，是社会建设的兜底性、基础性工作。各级党委和政府要坚持以人民为中心，加强对民政工作的领导，增强基层民政服务能力，推动民政事业持续健康发展。各级民政部门要加强党的建设，坚持改革创新，聚焦脱贫攻坚，聚焦特殊群体，聚焦群众关切，更好履行基本民生保障、基层社会治理、基本社会服务等职责，为全面建成小康社会、全面建设社会主义现代化国家作出新的贡献。\n" +
                "\n" +
                "　　第十四次全国民政会议4月2日在北京召开。会上传达了习近平的重要指示。");
        System.out.println("====== 初始化信息 ======");
        return map;
    }


    /**
     * 设置成绩信息
     *
     * @return
     */
    public static Map<String, String> getScoreData() {
        Map<String, Map<String, String>> map = new TreeMap();
        Map<String, String> map1 = new HashMap();
        for (int i = 1; i < count1; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx1" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx1" + i, "公共必修");
            map1.put("xf1" + i, "2");
            map1.put("cj1" + i, "65");
            //            map.put("2014-2015学年第一学期", map1);
            map.put("name1" + i, map1);

        }
        for (int i = 1; i < count2; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx2" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx2" + i, "公共必修");
            map1.put("xf2" + i, "2");
            map1.put("cj2" + i, "65");
            map.put("name2" + i, map1);

        }
        for (int i = 1; i < count3; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx3" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx3" + i, "公共必修");
            map1.put("xf3" + i, "2");
            map1.put("cj3" + i, "65");
            //            map.put("2015-2016学年第一学期", map1);
            map.put("name3" + i, map1);

        }
        for (int i = 1; i < count4; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx4" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx4" + i, "公共必修");
            map1.put("xf4" + i, "2");
            map1.put("cj4" + i, "65");
            //            map.put("2015-2016学年第二学期", map1);
            map.put("name4" + i, map1);
        }
        for (int i = 1; i < count5; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx5" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx5" + i, "公共必修");
            map1.put("xf5" + i, "2");
            map1.put("cj5" + i, "65");
            //            map.put("2016-2017学年第一学期", map1);
            map.put("name5" + i, map1);
        }
        for (int i = 1; i < count6; i++) {
            //            Map<String, String> map1 = new HashMap<>();
            map1.put("kcmx6" + i, "职业与创业教育就业指导(一)" + i);
            map1.put("kcsx6" + i, "公共必修");
            map1.put("xf6" + i, "2");
            map1.put("cj6" + i, "65");
            map.put("name6" + i, map1);
            //            map.put("2016-2017学年第二学期", map1);
        }
        return map1;
    }


}
