package com.weixf;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.AttachmentType;
import com.deepoove.poi.data.Attachments;
import com.deepoove.poi.data.Charts;
import com.deepoove.poi.data.DocumentRenderData;
import com.deepoove.poi.data.Documents;
import com.deepoove.poi.data.HyperlinkTextRenderData;
import com.deepoove.poi.data.Includes;
import com.deepoove.poi.data.MergeCellRule;
import com.deepoove.poi.data.NumberingFormat;
import com.deepoove.poi.data.Numberings;
import com.deepoove.poi.data.Paragraphs;
import com.deepoove.poi.data.PictureType;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.Rows;
import com.deepoove.poi.data.Tables;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.Texts;
import com.deepoove.poi.data.style.BorderStyle;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.plugin.comment.CommentRenderData;
import com.deepoove.poi.plugin.comment.CommentRenderPolicy;
import com.deepoove.poi.plugin.comment.Comments;
import com.deepoove.poi.plugin.highlight.HighlightRenderData;
import com.deepoove.poi.plugin.highlight.HighlightRenderPolicy;
import com.deepoove.poi.plugin.highlight.HighlightStyle;
import com.deepoove.poi.plugin.markdown.MarkdownRenderData;
import com.deepoove.poi.plugin.markdown.MarkdownRenderPolicy;
import com.deepoove.poi.plugin.markdown.MarkdownStyle;
import com.deepoove.poi.plugin.table.LoopColumnTableRenderPolicy;
import com.deepoove.poi.plugin.table.LoopRowTableRenderPolicy;
import com.deepoove.poi.policy.AttachmentRenderPolicy;
import com.deepoove.poi.util.BufferedImageUtils;
import com.documents4j.api.DocumentType;
import com.documents4j.api.IConverter;
import com.documents4j.job.LocalConverter;
import com.weixf.model.AddrModel;
import com.weixf.model.DetailData;
import com.weixf.model.Dog;
import com.weixf.model.Goods;
import com.weixf.model.Labor;
import com.weixf.model.LoopInnerVar;
import com.weixf.model.LoopPlugin;
import com.weixf.plugin.custom.CustomDetailTablePolicy;
import com.weixf.plugin.custom.CustomImgRenderPolicy;
import com.weixf.plugin.custom.CustomListRenderPolicy;
import com.weixf.plugin.custom.CustomMyTablePolicy;
import com.weixf.plugin.custom.CustomTextRenderPolicy;
import org.apache.poi.util.LocaleUtil;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.awt.*;
import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileInputStream;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Base64;
import java.util.Date;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Random;
import java.util.concurrent.TimeUnit;

/**
 * @author weixf
 */
@DisplayName("模板测试")
public class PoitlTest {

    public static String base;
    public static String templatePath;
    public static String imgPath;
    public static String markdownPath;
    public static String outPath;
    public BufferedImage bufferImage;
    public String imageBase64;


    public void config() {
        Properties properties = System.getProperties();
        String projectPath = (String) properties.get("user.dir");
        base = projectPath + "/src/test/resources/";
        templatePath = base + "template/";
        // templatePath = base + "templates/";
        imgPath = base + "images/";
        markdownPath = base + "markdown/";
        outPath = base + "out/";
    }

    public void createImages() throws Exception {
        bufferImage = BufferedImageUtils.newBufferImage(100, 100);
        Graphics2D g = (Graphics2D) bufferImage.getGraphics();
        g.setColor(Color.CYAN);
        g.fillRect(0, 0, 100, 100);
        g.setColor(Color.BLACK);
        g.drawString("Java Image", 0, 50);
        g.dispose();
        bufferImage.flush();

        File imageFile = new File(base + "fruits.jpg");
        FileInputStream fis = new FileInputStream(imageFile);
        byte[] fileBytes = new byte[(int) imageFile.length()];
        int read = fis.read(fileBytes);
        fis.close();
        String imageBase64String = Base64.getEncoder().encodeToString(fileBytes);
        imageBase64 = "data:image/png;base64," + imageBase64String;
    }

    @BeforeEach
    public void init() throws Exception {
        config();
        createImages();
    }

    // @AfterEach
    public void testPdf() {
        File wordFile = new File(outPath + "out_MyTemplate.docx");
        File targetPdf = new File(outPath + "out_MyTemplate.pdf");
        IConverter converter = LocalConverter.builder()
                .baseFolder(new File(outPath))
                .workerPool(5, 8, 2, TimeUnit.SECONDS)
                .processTimeout(5, TimeUnit.SECONDS)
                .build();
        boolean execute = converter.convert(wordFile)
                .as(DocumentType.DOCX)
                .to(targetPdf)
                .as(DocumentType.PDF)
                .execute();
        if (execute) {
            converter.shutDown();
        } else {
            converter.kill();
        }
    }

    @Test
    public void testRenderTemplate() throws Exception {

        Map<String, Object> data = new HashMap<>();

        // 一级标题
        data.put("title", "Poi-tl Documentation");
        // 作者
        data.put("author", "Sayi");
        // 邮箱
        data.put("email", Texts.of("sayi@apache.org")
                .mailto("sayi@apache.org", "poi-tl")
                .fontSize(10)
                .create());
        // 版本
        data.put("version", "Version 1.12.0");
        // 摘要
        data.put("abstract", "poi-tl（poi template language）是Word模板引擎，使用Word模板和数据创建很棒的Word文档");
        // 目标
        data.put("target", Texts.of("在文档的任何地方做任何事情（Do Anything Anywhere）是poi-tl的星辰大海。")
                .italic()
                .create());
        // 开源协议
        data.put("license", "Apache License 2.0");
        // 源码地址
        data.put("sourceLink", Texts.of("GitHub").link("https://github.com/Sayi/poi-tl").create());
        // 为什么选择poi-tl
        data.put("solutionCompare", Tables.create(
                Rows.of("方案", "移植性", "功能性", "易用性").textColor("FFFFFF").bgColor("4472C4").textBold().textFontFamily("宋体").textFontSize(12).create(),
                Rows.create("Poi-tl", "Java跨平台", "Word模板引擎，基于Apache POI，提供更友好的API", "低代码，准备文档模板和数据即可"),
                Rows.create("Apache POI", "Java跨平台", "Apache项目，封装了常见的文档操作，也可以操作底层XML结构", "-"),
                Rows.create("Freemarker", "XML跨平台", "仅支持文本，很大的局限性", "仅支持文本，很大的局限性"),
                Rows.create("OpenOffice", "部署OpenOffice，移植性较差", "-", "需要了解OpenOffice的API"),
                Rows.create("HTML浏览器导出", "依赖浏览器的实现，移植性较差", "HTML不能很好的兼容Word的格式，样式糟糕", "-"),
                Rows.create("Jacob、winlib", "Windows平台", "-", "复杂，完全不推荐使用")));
        // poi-tl功能列举
        data.put("funcDesc", Tables.create(
                Rows.of("Word模板引擎功能", "描述").textColor("FFFFFF").bgColor("4472C4").textBold().textFontFamily("宋体").textFontSize(12).create(),
                Rows.create("文本", "将标签渲染为文本"),
                Rows.create("图片", "将标签渲染为图片"),
                Rows.create("表格", "将标签渲染为表格"),
                Rows.create("列表", "将标签渲染为列表"),
                Rows.create("图表", "条形图（3D条形图）、柱形图（3D柱形图）、面积图（3D面积图）、折线图（3D折线图）、雷达图、饼图（3D饼图）、散点图等图表渲染"),
                Rows.create("If Condition判断", "根据条件隐藏或者显示某些文档内容（包括文本、段落、图片、表格、列表、图表等）"),
                Rows.create("Foreach Loop循环", "根据集合循环某些文档内容（包括文本、段落、图片、表格、列表、图表等）"),
                Rows.create("Loop表格行", "循环复制渲染表格的某一行"),
                Rows.create("Loop表格列", "循环复制渲染表格的某一列"),
                Rows.create("Loop有序列表", "支持有序列表的循环，同时支持多级列表"),
                Rows.create("Highlight代码高亮", "word中代码块高亮展示，支持26种语言和上百种着色样式"),
                Rows.create("Markdown", "将Markdown渲染为word文档"),
                Rows.create("Word批注", "完整的批注功能，创建批注、修改批注等"),
                Rows.create("Word附件", "Word中插入附件"),
                Rows.create("SDT内容控件", "Word中插入附件"),
                Rows.create("Textbox文本框", "Word中插入附件"),
                Rows.create("图片替换", "图片替换"),
                Rows.create("书签、锚点、超链接", "支持设置书签，文档内锚点和超链接功能"),
                Rows.create("Expression Language", "完全支持SpringEL表达式，可以扩展更多的表达式：OGNL, MVEL…"),
                Rows.create("样式", "模板即样式，同时代码也可以设置样式"),
                Rows.create("模板嵌套", "模板包含子模板，子模板再包含子模板"),
                Rows.create("合并", "Word合并Merge，也可以在指定位置进行合并"),
                Rows.create("用户自定义函数(插件)", "插件化设计，在文档任何位置执行函数")
        ));
        // maven依赖
        HighlightRenderData MavenImportCode = new HighlightRenderData();
        MavenImportCode.setCode("<dependency>\n" +
                "  <groupId>com.deepoove</groupId>\n" +
                "  <artifactId>poi-tl</artifactId>\n" +
                "  <version>1.12.0</version>\n" +
                "</dependency>");
        MavenImportCode.setLanguage("java");
        MavenImportCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("mavenImportCode", MavenImportCode);
        // 快速入门
        HighlightRenderData SimpleCode = new HighlightRenderData();
        SimpleCode.setCode("XWPFTemplate template = XWPFTemplate.compile(\"template.docx\").render(\n" +
                "  new HashMap<String, Object>(){{\n" +
                "    put(\"title\", \"Hi, poi-tl Word模板引擎\");\n" +
                "}});\n" +
                "template.writeAndClose(new FileOutputStream(\"output.docx\"));");
        SimpleCode.setLanguage("java");
        SimpleCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleCode", SimpleCode);
        // 标签替换成标签文本
        data.put("label1", "{{title}}");
        data.put("label2", "{{开头，以}}");
        data.put("label3", "{{author.name}}");
        data.put("label4", "{{path}}");
        data.put("label5", "{{width}}");
        data.put("label6", "{{height}}");
        data.put("label7", "{{var}}");
        data.put("label8", "{{@var}}");
        data.put("label9", "{{#var}}");
        data.put("label10", "{{#var}}");
        data.put("label11", "{{?sections}}{{/sections}}");
        data.put("label12", "{\n" +
                "  \"announce\": false\n" +
                "}\n" +
                "--------------\n" +
                "Made it,Ma!{{?announce}}Top of the world!{{/announce}}\n" +
                "Made it,Ma!\n{{?announce}}\nTop of the world!\uD83C\uDF8B\n{{/announce}}");
        data.put("label13", "{\n" +
                "  \"person\": { \"name\": \"Sayi\" }\n" +
                "}\n" +
                "--------------\n" +
                "{{?person}}\nHi {{name}}!\n{{/person}}");
        data.put("label14", "{\n" +
                "  \"songs\": [\n" +
                "    { \"name\": \"Memories\" },\n" +
                "    { \"name\": \"Sugar\" },\n" +
                "    { \"name\": \"Last Dance\" }\n" +
                "  ]\n" +
                "}\n" +
                "--------------\n" +
                "{{?songs}}\n{{name}}\n{{/songs}}");
        data.put("label15", "{\n" +
                "  \"produces\": [\n" +
                "    \"application/json\",\n" +
                "    \"application/xml\"\n" +
                "  ]\n" +
                "}\n" +
                "----------------\n" +
                "{{?produces}}\n{{_index + 1}}. {{=#this}}\n{{/produces}}");
        data.put("label16", "{ {+var} }");
        data.put("label17", "{ {+nested} }");
        data.put("label18", "{{addr}}");
        data.put("label19", "{{var}}");
        data.put("label20", "{{}}");
        data.put("label21", "{{name}}\n\n" +
                "类方法调用，转大写：{{name.toUpperCase()}}\n\n" +
                "判断条件：{{name == 'poi-tl'}}\n\n" +
                "{{empty?:'这个字段为空'}}\n\n" +
                "三目运算符：{{sex ? '男' : '女'}}\n\n" +
                "类方法调用，时间格式化：{{new java.text.SimpleDateFormat('yyyy-MM-dd HH:mm:ss').format(time)}}\n\n" +
                "运算符：{{price/10000 + '万元'}}\n\n" +
                "数组列表使用下标访问：{{dogs[0].name}}\n\n" +
                "使用静态类方法：{{localDate.format(T(java.time.format.DateTimeFormatter).ofPattern('yyyy年MM月dd日'))}}");
        data.put("label22", "{{?desc == null or desc == ''}}{{summary}}{{/}}\n\n" +
                "{{?produces == null or produces.size() == 0}}无{{/}}");
        data.put("label23", "{ {/} }");
        data.put("label24", "{{report}}");
        data.put("label25", "{{%var}}");
        data.put("label26", "{{goods}}");
        data.put("label27", "{{labors}}");
        data.put("label28", "{{detail_table}}");


        // 数据模型 - Map
        HighlightRenderData MapDataCode = new HighlightRenderData();
        MapDataCode.setCode("Map<String, Object> data = new HashMap<>();\n" +
                "data.put(\"name\", \"Sayi\");\n" +
                "data.put(\"start_time\", \"2019-08-04\");");
        MapDataCode.setLanguage("java");
        MapDataCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("monokai").build());
        data.put("mapDataCode", MapDataCode);
        // 数据模型 - 对象
        HighlightRenderData ObjectDataCode = new HighlightRenderData();
        ObjectDataCode.setCode("public class Data {\n" +
                "  private String name;\n" +
                "  private String startTime;\n" +
                "  private Author author;\n" +
                "}");
        ObjectDataCode.setLanguage("java");
        ObjectDataCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("monokai").build());
        data.put("objectDataCode", ObjectDataCode);
        // 输出示例
        HighlightRenderData OutputCode = new HighlightRenderData();
        OutputCode.setCode("response.setContentType(\"application/octet-stream\");\n" +
                "response.setHeader(\"Content-disposition\",\"attachment;filename=\\\"\"+\"out_template.docx\"+\"\\\"\");\n" +
                "\n" +
                "// HttpServletResponse response\n" +
                "OutputStream out = response.getOutputStream();\n" +
                "BufferedOutputStream bos = new BufferedOutputStream(out);\n" +
                "template.write(bos);\n" +
                "bos.flush();\n" +
                "out.flush();\n" +
                "PoitlIOUtils.closeQuietlyMulti(template, bos, out);");
        OutputCode.setLanguage("java");
        OutputCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("ocean").build());
        data.put("outputCode", OutputCode);

        // 文本标签 {{var}}
        HighlightRenderData simpleTextRenderCode = new HighlightRenderData();
        simpleTextRenderCode.setCode("put(\"name\", \"Sayi\");\n" +
                "put(\"author\", new TextRenderData(\"000000\", \"Sayi\"));\n" +
                "put(\"link\", new HyperlinkTextRenderData(\"website\", \"http://deepoove.com\"));\n" +
                "put(\"anchor\", new HyperlinkTextRenderData(\"anchortxt\", \"anchor:appendix1\"));");
        simpleTextRenderCode.setLanguage("java");
        simpleTextRenderCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleTextRenderCode", simpleTextRenderCode);

        HighlightRenderData simpleChainTextRenderCode = new HighlightRenderData();
        simpleChainTextRenderCode.setCode("put(\"author\", Texts.of(\"Sayi\").color(\"000000\").create());\n" +
                "put(\"link\", Texts.of(\"website\").link(\"http://deepoove.com\").create());\n" +
                "put(\"anchor\", Texts.of(\"anchortxt\").anchor(\"appendix1\").create());");
        simpleChainTextRenderCode.setLanguage("java");
        simpleChainTextRenderCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleChainTextRenderCode", simpleChainTextRenderCode);

        data.put("showSimpleTextRenderData",
                Includes.ofLocal(base + "simpleTextRenderDataTemplate.docx")
                        .setRenderModel(new HashMap<String, Object>() {
                            {
                                put("name", "Sayi");
                                put("author1", new TextRenderData("28a745", "我是\n绿色且换行的文字"));
                                put("link1", new HyperlinkTextRenderData("poi-tl网站", "http://deepoove.com"));
                                put("anchor1", new HyperlinkTextRenderData("anchortxt", "anchor:Sayi"));
                                put("author2", Texts.of("我是绿色且换行的\n文字").color("28a745").create());
                                put("link2", Texts.of("poi-tl网站").link("http://deepoove.com").create());
                                put("anchor2", Texts.of("anchortxt").anchor("Sayi").create());
                            }
                        }).create());

        // 图片标签 {{@var}}
        HighlightRenderData simplePicturesRenderCode = new HighlightRenderData();
        simplePicturesRenderCode.setCode("// 指定图片路径\n" +
                "put(\"localPicture\", Pictures.ofLocal(\"sayi.png\").size(120, 120).create());\n" +
                "// 设置图片宽高\n" +
                "put(\"imageByWidthAndHeight\", Pictures.ofLocal(\"logo.png\").size(120, 120).create());\n" +
                "// 图片流\n" +
                "put(\"localBytePicture\", Pictures.ofStream(Files.newInputStream(Paths.get(\"logo.png\"))).size(100, 120).create());\n" +
                "// 网络图片\n" +
                "put(\"urlPicture\", \"http://deepoove.com/images/icecream.png\");\n" +
                "// java图片\n" +
                "put(\"bufferImagePicture\", Pictures.ofBufferedImage(bufferImage, PictureType.PNG).size(100, 100).create());\n" +
                "// base64图片\n" +
                "put(\"base64Image\", Pictures.ofBase64(imageBase64, PictureType.PNG).size(100, 100).center().create());\n" +
                "// svg图片\n" +
                "put(\"svgPicture\", Pictures.ofUrl(\"http://deepoove.com/images/%E8%8C%84%E5%AD%90.svg\").create());");
        simplePicturesRenderCode.setLanguage("java");
        simplePicturesRenderCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simplePicturesRenderCode", simplePicturesRenderCode);

        data.put("showSimplePicturesRenderData",
                Includes.ofLocal(base + "simplePicturesRenderDataTemplate.docx").setRenderModel(
                        new HashMap<String, Object>() {
                            {
                                put("localPicture", Pictures.ofLocal(base + "sayi.png").size(120, 120).create());
                                put("imageByWidthAndHeight", Pictures.ofLocal(base + "logo.png").size(220, 220).create());
                                put("localBytePicture", Pictures.ofStream(Files.newInputStream(Paths.get(base + "logo.png"))).size(100, 120).create());
                                put("urlPicture", "http://deepoove.com/images/icecream.png");
                                put("bufferImagePicture", Pictures.ofBufferedImage(bufferImage, PictureType.PNG).size(100, 100).right().create());
                                put("base64Image", Pictures.ofBase64(imageBase64, PictureType.PNG).size(100, 100).center().create());
                                put("svgPicture", Pictures.ofUrl("http://deepoove.com/images/%E8%8C%84%E5%AD%90.svg").size(50, 50).create());
                            }
                        }).create());

        // 表格标签 {{#var}}
        HighlightRenderData simpleTableCode1 = new HighlightRenderData();
        simpleTableCode1.setCode("// 一个2行2列的表格\n" +
                "put(\"table0\", Tables.of(new String[][] {\n" +
                "                new String[] { \"00\", \"01\" },\n" +
                "                new String[] { \"10\", \"11\" }\n" +
                "            }).border(BorderStyle.DEFAULT).create());");
        simpleTableCode1.setLanguage("java");
        simpleTableCode1.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleTableCode1", simpleTableCode1);

        HighlightRenderData simpleTableCode2 = new HighlightRenderData();
        simpleTableCode2.setCode("// 第0行居中且背景为蓝色的表格\n" +
                "RowRenderData row0 = Rows.of(\"姓名\", \"学历\").textColor(\"FFFFFF\")\n" +
                "      .bgColor(\"4472C4\").center().create();\n" +
                "RowRenderData row1 = Rows.create(\"李四\", \"博士\");\n" +
                "put(\"table1\", Tables.create(row0, row1));");
        simpleTableCode2.setLanguage("java");
        simpleTableCode2.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleTableCode2", simpleTableCode2);

        HighlightRenderData simpleMergeTableCode = new HighlightRenderData();
        simpleMergeTableCode.setCode("// 合并第1行所有单元格的表格\n" +
                "RowRenderData row0 = Rows.of(\"列0\", \"列1\", \"列2\").center().bgColor(\"4472C4\").create();\n" +
                "RowRenderData row1 = Rows.create(\"没有数据\", null, null);\n" +
                "MergeCellRule rule = MergeCellRule.builder().map(Grid.of(1, 0), Grid.of(1, 2)).build();\n" +
                "put(\"table3\", Tables.of(row0, row1).mergeRule(rule).create());");
        simpleMergeTableCode.setLanguage("java");
        simpleMergeTableCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleMergeTableCode", simpleMergeTableCode);

        data.put("showSimpleTableRenderData", Includes.ofLocal(base + "simpleTableRenderDataTemplate.docx").setRenderModel(
                new HashMap<String, Object>() {
                    {
                        put("table0", Tables.of(new String[][]{new String[]{"00", "01"}, new String[]{"10", "11"}}).border(BorderStyle.DEFAULT).create());
                        put("table1", Tables.create(
                                Rows.of("姓名", "学历").textColor("FFFFFF").bgColor("4472C4").center().create(),
                                Rows.create("李四", "博士")
                        ));
                        put("table2", Tables.of(
                                        Rows.of("列0", "列1", "列2").center().bgColor("4472C4").create(),
                                        Rows.create("没有数据", null, null))
                                .mergeRule(MergeCellRule.builder()
                                        .map(MergeCellRule.Grid.of(1, 0), MergeCellRule.Grid.of(1, 2))
                                        .build())
                                .create()
                        );
                    }
                }).create());

        // 列表标签 {{*var}}
        HighlightRenderData simpleListRenderCode = new HighlightRenderData();
        simpleListRenderCode.setCode("put(\"list\", Numberings.of(NumberingFormat.DECIMAL)\n" +
                " .addItem(\"Plug-in grammar\")\n" +
                " .addItem(\"Supports word text, pictures, table...\")\n" +
                " .addItem(\"Not just templates\")\n" +
                " .addItem(\"啊啊啊啊啊啊啊啊啊啊啊啊阿啊啊啊\")\n" +
                " .create());");
        simpleListRenderCode.setLanguage("java");
        simpleListRenderCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleListRenderCode", simpleListRenderCode);

        data.put("showSimpleListRenderData", Includes.ofLocal(base + "simpleListRenderDataTemplate.docx")
                .setRenderModel(new HashMap<String, Object>() {
                    {
                        put("list", Numberings.of(NumberingFormat.DECIMAL)
                                .addItem("Plug-in grammar")
                                .addItem(Texts.of("Supports word text, pictures, table...").color("FF0000").create())
                                .addItem("Not just templates")
                                .addItem("啊啊啊啊啊啊啊啊啊啊啊啊阿啊啊啊")
                                .create());
                    }
                }).create());


        // 区块对标签 {{?sections}}{{/sections}}
        data.put("announce", false);
        data.put("person", new HashMap<String, Object>() {
            {
                put("name", "Sayi");
            }
        });
        data.put("songs", new ArrayList<Map<String, Object>>() {
            {
                add(new HashMap<String, Object>() {
                    {
                        put("name", "Memories");
                    }
                });
                add(new HashMap<String, Object>() {
                    {
                        put("name", "Sugar");
                    }
                });
                add(new HashMap<String, Object>() {
                    {
                        put("name", "Last Dance");
                    }
                });
            }
        });
        data.put("produces", new ArrayList<String>() {{
            add("application/json");
            add("application/xml");
            add("aaaaaaaaaaaaaaaa");
            add("bbbbbbbbbbbbbbbb");
        }});
        // 区块对标签内置循环变量
        data.put("loopInnerVars", new ArrayList<LoopInnerVar>() {{
            add(new LoopInnerVar("_index", "int", "返回当前迭代从0开始的索引"));
            add(new LoopInnerVar("_is_first", "boolean", "辨别循环项是否是当前迭代的第一项"));
            add(new LoopInnerVar("_is_last", "boolean", "辨别循环项是否是当前迭代的最后一项"));
            add(new LoopInnerVar("_has_next", "boolean", "辨别循环项是否是有下一项"));
            add(new LoopInnerVar("_is_even_item", "boolean", "辨别循环项是否是当前迭代间隔1的奇数项"));
            add(new LoopInnerVar("_is_odd_item", "boolean", "辨别循环项是否是当前迭代间隔1的偶数项"));
            add(new LoopInnerVar("#this", "object", "引用当前对象，由于#和已有表格标签标识冲突，所以在文本标签中需要使用=号标识来输出文本"));
        }});

        // 嵌套标签 {{+var}}
        HighlightRenderData simpleNestRenderCode = new HighlightRenderData();
        simpleNestRenderCode.setCode("class AddrModel {\n" +
                "  private String addr;\n" +
                "  public AddrModel(String addr) {\n" +
                "    this.addr = addr;\n" +
                "  }\n" +
                "  // Getter/Setter\n" +
                "}\n" +
                "\n" +
                "List<AddrModel> subData = new ArrayList<>();\n" +
                "subData.add(new AddrModel(\"Hangzhou,China\"));\n" +
                "subData.add(new AddrModel(\"Shanghai,China\"));\n" +
                "put(\"showSimpleNestRenderData\", Includes.ofLocal(\"showSimpleNestRenderDataTemplate.docx\").setRenderModel(subData).create());");
        simpleNestRenderCode.setLanguage("java");
        simpleNestRenderCode.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("zenburn").build());
        data.put("simpleNestRenderCode", simpleNestRenderCode);

        data.put("showSimpleNestRenderData", Includes.ofLocal(base + "showSimpleNestRenderDataTemplate.docx").setRenderModel(
                new ArrayList<AddrModel>() {
                    {
                        add(new AddrModel("Hangzhou,China"));
                        add(new AddrModel("Shanghai,China"));
                    }
                }).create());


        // 引用标签 {{var}} 设置图片格式—可选文字—标题或者说明 or 编辑替换文字-替换文字
        data.put("refPicture", Pictures.ofLocal(base + "pictureRef.png").create());
        data.put("refPic", Pictures.ofLocal(base + "sayi.png").create());
        // 多系列图表
        data.put("refPicMultiSeries", Pictures.ofLocal(base + "pictureRefMultiSeries.png").create());
        data.put("barChart", Charts
                .ofMultiSeries("CustomTitle",
                        new String[]{"中文", "English", "日本語", "português", "中文", "English", "日本語", "português"})
                .addSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0, 150.0, 6.0, 118.0, 31.0})
                .addSeries("speakers", new Double[]{223.0, 119.0, 154.0, 142.0, 223.0, 119.0, 54.0, 42.0})
                .addSeries("youngs", new Double[]{323.0, 89.0, 54.0, 42.0, 223.0, 119.0, 54.0, 442.0})
                .create());
        // 单系列图表
        data.put("refPicSingleSeries", Pictures.ofLocal(base + "pictureRefSingleSeries.png").create());
        data.put("pieChart", Charts.ofSingleSeries("ChartTitle",
                        new String[]{"俄罗斯", "加拿大", "美国", "中国", "巴西", "澳大利亚", "印度"})
                .series("countries", new Integer[]{17098242, 9984670, 9826675, 9596961, 8514877, 7741220, 3287263})
                .create());
        // 组合图表
        data.put("pictureRefCombo", Pictures.ofLocal(base + "pictureRefCombo.jpeg").create());
        data.put("combChart", Charts.ofComboSeries("ComboChartTitle",
                        new String[]{"中文", "English", "日本語", "português"})
                .addBarSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0})
                .addBarSeries("speakers", new Double[]{223.0, 119.0, 154.0, 142.0})
                .addBarSeries("NewBar", new Double[]{223.0, 119.0, 154.0, 142.0})
                .addLineSeries("youngs", new Double[]{323.0, 89.0, 54.0, 42.0})
                .addLineSeries("NewLine", new Double[]{123.0, 59.0, 154.0, 42.0})
                .create());

        // Spring表达式
        data.put("name", "poi-tl");
        data.put("empty", "");
        data.put("sex", true);
        data.put("time", new Date());
        data.put("price", 100000);
        Dog[] dogArray = new Dog[10];
        for (int i = 0; i < 10; i++) {
            Dog dog = new Dog();
            dog.setName("二哈-" + i);
            dog.setAge(i + 100);
            dogArray[i] = dog;
        }
        data.put("dogs", dogArray);
        data.put("localDate", LocalDate.now());

        data.put("desc", "");
        data.put("summary", "Find A Pet");
        data.put("produces", new ArrayList<String>() {
            {
                add("application/xml");
                add("application/xml");
                add("application/xml");
                add("application/xml");
            }
        });

        // 默认插件列举
        data.put("pluginListDefault", Numberings.of(NumberingFormat.DECIMAL_PARENTHESES)
                .addItem("TextRenderPolicy")
                .addItem("PictureRenderPolicy")
                .addItem("NumberingRenderPolicy")
                .addItem("TableRenderPolicy")
                .addItem("DocxRenderPolicy")
                .addItem("MultiSeriesChartTemplateRenderPolicy")
                .addItem("SingleSeriesChartTemplateRenderPolicy")
                .addItem("DefaultPictureTemplateRenderPolicy")
                .create());

        // 自定义插件 CustomTextRenderPolicy、CustomImgRenderPolicy、CustomListRenderPolicy、CustomMyTablePolicy
        data.put("sea", "Hello, world!");
        data.put("sea_img", base + "aaa.png");
        data.put("sea_feature", Arrays.asList("面朝大海春暖花开", "今朝有酒今朝醉"));
        data.put("sea_location", Arrays.asList("日落：日落山花红四海", "花海：你想要的都在这里"));

        // 更多插件列举
        data.put("loopPlugins", new ArrayList<LoopPlugin>() {
            {
                add(new LoopPlugin("ParagraphRenderPolicy", "渲染一个段落，可以包含不同样式文本，图片等", null));
                add(new LoopPlugin("DocumentRenderPolicy", "渲染一个段落，可以包含不同样式文本，图片等", null));
                add(new LoopPlugin("CommentRenderPolicy", "完整的批注功能",
                        Texts.of("示例-批注").anchor("10.5. 批注").create()));
                add(new LoopPlugin("AttachmentRenderPolicy", "插入附件功能",
                        Texts.of("示例-插入附件").anchor("10.6. 插入附件").create()));
                add(new LoopPlugin("LoopRowTableRenderPolicy", "循环表格行",
                        Texts.of("示例-表格行循环").anchor("10.2. 表格行循环").create()));
                add(new LoopPlugin("LoopColumnTableRenderPolicy", "循环表格列",
                        Texts.of("示例-表格行循环").anchor("10.3. 表格列循环").create()));
                add(new LoopPlugin("DynamicTableRenderPolicy", "动态表格插件，允许直接操作表格对象",
                        Texts.of("示例-动态表格").anchor("10.4. 动态表格").create()));
                add(new LoopPlugin("BookmarkRenderPolicy", "书签和锚点",
                        Texts.of("示例-Swagger文档").link("http://deepoove.com/poi-tl/#example-swagger").create()));
                add(new LoopPlugin("AbstractChartTemplateRenderPolicy", "引用图表插件，允许直接操作图表对象", null));
                add(new LoopPlugin("TOCRenderPolicy", "Beta实验功能：目录，打开文档时会提示更新域", null));
                add(new LoopPlugin("HighlightRenderPolicy", "Word支持代码高亮",
                        Texts.of("示例-代码高亮").anchor("10.7. 代码高亮").create()));
                add(new LoopPlugin("MarkdownRenderPolicy", "使用Markdown来渲染word",
                        Texts.of("示例-Markdown").anchor("10.8. Markdown").create()));
            }
        });

        // 循环表格行
        data.put("goods", new ArrayList<Goods>() {
            {
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
            }
        });
        data.put("labors", new ArrayList<Labor>() {
            {
                add(new Labor("油漆工", 2, 400, 1600));
                add(new Labor("油漆工", 2, 400, 1600));
                add(new Labor("油漆工", 2, 400, 1600));
                add(new Labor("油漆工", 2, 400, 1600));
            }
        });
        data.put("total", 1024);

        // 循环表格列
        data.put("goods2", new ArrayList<Goods>() {
            {
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
                add(new Goods(4, "墙纸", "书房卧室", 1500, 400,
                        new Random().nextInt(10) + 20, 1600,
                        Pictures.ofLocal(base + "earth.png").size(24, 24).create()));
            }
        });

        // 动态表格
        DetailData detailData = new DetailData();
        RowRenderData good = Rows.of("4", "墙纸", "书房卧室", "1500", "/", "400", "1600").center().create();
        List<RowRenderData> goods = Arrays.asList(good, good, good);
        RowRenderData labor = Rows.of("油漆工", "2", "200", "400").center().create();
        List<RowRenderData> labors = Arrays.asList(labor, labor, labor, labor);
        detailData.setGoods(goods);
        detailData.setLabors(labors);
        data.put("detail_table", detailData);
        data.put("d_total", 200);

        // 插入批注
        data.put("poetryTitle", "Sayi");
        data.put("pic", Pictures.ofLocal(base + "sayi.png").size(60, 60).create());
        data.put("poetryAuthor", "骆宾王作为“初唐四杰”之一，对荡涤六朝文学颓波，革新初唐浮靡诗风。他一生著作颇丰，是一个才华横溢的诗人。");
        data.put("interpretation1", "弯着脖子");
        data.put("interpretation2", "划动");

        // 插入附件
        data.put("attachment", Attachments.ofLocal(templatePath + "attachment.docx", AttachmentType.DOCX).create());

        // 代码高亮
        HighlightRenderData highlightRenderData = new HighlightRenderData();
        highlightRenderData.setCode("/**\n"
                + " * @author John Smith <john.smith@example.com>\n"
                + "*/\n"
                + "package l2f.gameserver.model;\n"
                + "\n"
                + "public abstract strictfp class L2Char extends L2Object {\n"
                + "  public static final Short ERROR = 0x0001;\n"
                + "\n"
                + "  public void moveTo(int x, int y, int z) {\n"
                + "    _ai = null;\n"
                + "    log(\"Should not be called\");\n"
                + "    if (1 > 5) { // wtf!?\n"
                + "      return;\n"
                + "    }\n"
                + "  }\n"
                + "}");
        highlightRenderData.setLanguage("java");
        highlightRenderData.setStyle(HighlightStyle.builder().withShowLine(true).withTheme("github").build());
        data.put("code", highlightRenderData);

        // 通过Markdown生成word文档
        MarkdownRenderData markdownRenderData = new MarkdownRenderData();
        markdownRenderData.setMarkdown(new String(Files.readAllBytes(Paths.get(markdownPath + "数据结构与算法.md"))));
        markdownRenderData.setStyle(MarkdownStyle.newStyle());
        data.put("md", markdownRenderData);

        Configure config = Configure.builder()
                .useSpringEL()
                .bind("mavenImportCode", new HighlightRenderPolicy())
                .bind("simpleCode", new HighlightRenderPolicy())
                .bind("mapDataCode", new HighlightRenderPolicy())
                .bind("objectDataCode", new HighlightRenderPolicy())
                .bind("outputCode", new HighlightRenderPolicy())
                .bind("simpleTextRenderCode", new HighlightRenderPolicy())
                .bind("simpleChainTextRenderCode", new HighlightRenderPolicy())
                .bind("simplePicturesRenderCode", new HighlightRenderPolicy())
                .bind("simpleTableCode1", new HighlightRenderPolicy())
                .bind("simpleTableCode2", new HighlightRenderPolicy())
                .bind("simpleMergeTableCode", new HighlightRenderPolicy())
                .bind("simpleListRenderCode", new HighlightRenderPolicy())
                .bind("loopInnerVars", new LoopRowTableRenderPolicy())
                .bind("simpleNestRenderCode", new HighlightRenderPolicy())
                .bind("sea", new CustomTextRenderPolicy())
                .bind("sea_img", new CustomImgRenderPolicy())
                .bind("sea_feature", new CustomListRenderPolicy())
                .bind("sea_location", new CustomMyTablePolicy())
                .bind("loopPlugins", new LoopRowTableRenderPolicy())
                .bind("goods", new LoopRowTableRenderPolicy())
                .bind("labors", new LoopRowTableRenderPolicy())
                .bind("goods2", new LoopColumnTableRenderPolicy())
                .bind("detail_table", new CustomDetailTablePolicy())
                .bind("", new CommentRenderPolicy())
                .bind("attachment", new AttachmentRenderPolicy())
                .bind("code", new HighlightRenderPolicy())
                .bind("md", new MarkdownRenderPolicy())
                .build();

        XWPFTemplate.compile(base + "MyTemplate.docx", config)
                .render(data)
                .writeToFile(outPath + "out_MyTemplate.docx");


    }


    @Test
    public void testDocumentRender() throws Exception {
        // comment
        CommentRenderData comment0 = Comments.of()
                .signature("Sayi", "s", LocaleUtil.getLocaleCalendar())
                .addText(Texts.of("咏鹅").fontSize(20).bold().create())
                .comment(Documents.of()
                        .addParagraph(Paragraphs.of(Pictures.ofLocal(base + "logo.png").create()).create())
                        .create())
                .create();
        CommentRenderData comment1 = Comments.of()
                .signature("Sayi", "s", LocaleUtil.getLocaleCalendar())
                .addText("骆宾王")
                .comment("骆宾王作为“初唐四杰”之一，对荡涤六朝文学颓波，革新初唐浮靡诗风。他一生著作颇丰，是一个才华横溢的诗人。")
                .create();
        CommentRenderData comment2 = Comments.of()
                .signature("Sayi", "s", LocaleUtil.getLocaleCalendar())
                .addText("曲项").comment("弯着脖子")
                .create();
        CommentRenderData comment3 = Comments.of()
                .signature("Sayi", "s", LocaleUtil.getLocaleCalendar())
                .addText("拨").comment("划动")
                .create();
        // document
        DocumentRenderData documentRenderData = Documents.of()
                .addParagraph(Paragraphs.of().addComment(comment0).center().create())
                .addParagraph(Paragraphs.of().addComment(comment1).center().create())
                .addParagraph(Paragraphs.of("鹅，鹅，鹅，").addComment(comment2).addText("向天歌。").center().create())
                .addParagraph(Paragraphs.of("白毛浮绿水，红掌").addComment(comment3).addText("清波。").center().create())
                .create();
        Style style = Style.builder()
                .buildFontFamily("微软雅黑")
                .buildFontSize(14f)
                .build();
        XWPFTemplate.create(documentRenderData, style).writeToFile(outPath + "out_DocumentRender.docx");
    }
}
