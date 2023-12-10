package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.FilePictureRenderData;
import com.deepoove.poi.data.Includes;
import com.weixf.source.DataTest;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.nio.file.Files;
import java.util.Arrays;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

@DisplayName("Include Docx Render test case")
public class DocxRenderTest {

    public static String base;
    public static String templatePath;
    public static String imgPath;
    public static String markdownPath;
    public static String outPath;

    List<DataTest> dataList;

    @BeforeEach
    public void init() {

        Properties properties = System.getProperties();
        String projectPath = (String) properties.get("user.dir");
        base = projectPath + "/src/test/resources/";
        templatePath = base + "template/";
        // templatePath = base + "templates/";
        imgPath = base + "images/";
        markdownPath = base + "markdown/";
        outPath = base + "out/";


        DataTest data1 = new DataTest();
        data1.setQuestion("贞观之治是历史上哪个朝代");
        data1.setA("宋");
        data1.setB("唐");
        data1.setC("明");
        data1.setD("清");
        data1.setLogo(new FilePictureRenderData(120, 120, base + "sayi.png"));

        DataTest data2 = new DataTest();
        data2.setQuestion("康乾盛世是历史上哪个朝代");
        data2.setA("三国");
        data2.setB("元");
        data2.setC("清");
        data2.setD("唐");
        data2.setLogo(new FilePictureRenderData(100, 120, base + "logo.png"));

        dataList = Arrays.asList(data1, data2);
    }

    @Test
    public void testIncludeRender() throws Exception {

        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("title", "Hello, poi tl.");
                put("docx_template", Includes.ofLocal(templatePath + "render_include_merge_template.docx")
                        .setRenderModel(dataList).create());
                put("docx_template2", Includes.ofLocal(templatePath + "render_include_picture.docx").create());
                put("docx_template3", Includes.ofLocal(templatePath + "render_include_table.docx").create());
                put("docx_template4", Includes.ofStream(
                        Files.newInputStream(new File(templatePath + "render_include_all.docx").toPath())).create());
                put("newline", "Why poi-tl?");
            }
        };

        XWPFTemplate.compile(templatePath + "render_include.docx")
                .render(data)
                .writeToFile(outPath + "out_render_include.docx");

    }

}
