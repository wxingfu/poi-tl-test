package com.weixf.policy.ref;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.Pictures;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

@DisplayName("PictImage Render test case")
public class PictImageTemplateRenderTest {

    public static String base;
    public static String templatePath;
    public static String imgPath;
    public static String markdownPath;
    public static String outPath;

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
    }

    @Test
    public void testIterablePictRender() throws Exception {
        PictureRenderData image = Pictures.ofLocal(base + "sayi.png").size(100, 120).create();
        PictureRenderData image2 = Pictures.ofLocal(base + "logo.png").size(100, 120).create();
        Map<String, Object> data = new HashMap<String, Object>();
        data.put("title", "poi-tl");
        data.put("test_img", image);
        Map<String, Object> data1 = new HashMap<String, Object>();
        data1.put("title", "poi-tl2");
        data1.put("test_img", image2);

        List<Map<String, Object>> mores = new ArrayList<Map<String, Object>>();
        mores.add(data);
        mores.add(data1);

        XWPFTemplate.compile(templatePath + "reference_pict_iterable.docx")
                .render(new HashMap<String, Object>() {
                    {
                        put("mores", mores);
                    }
                })
                .writeToFile(outPath + "out_reference_pict_iterable.docx");
    }


    @Test
    public void testReplacePict() throws Exception {
        XWPFTemplate.compile(templatePath + "reference_pict.docx")
                .render(new HashMap<String, Object>() {
                    {
                        put("test_img", Pictures.ofLocal(base + "sayi.png").size(100, 120).create());
                    }
                }).writeToFile(outPath + "out_reference_pict.docx");
    }

}
