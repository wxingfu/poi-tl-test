package com.weixf.policy.ref;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.Pictures;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.*;

@DisplayName("Template Render test case")
public class PictureTemplateRenderTest {

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
    public void testIterablePictureRender() throws Exception {
        PictureRenderData image = Pictures.ofLocal(base + "sayi.png").size(100, 120).create();
        PictureRenderData image2 = Pictures.ofLocal(base + "logo.png").size(100, 120).create();
        Map<String, Object> data = new HashMap<String, Object>();
        data.put("title", "poi-tl");
        data.put("pic", image);
        Map<String, Object> data1 = new HashMap<String, Object>();
        data1.put("title", "poi-tl2");
        data1.put("pic", image2);

        List<Map<String, Object>> mores = new ArrayList<Map<String, Object>>();
        mores.add(data);
        mores.add(data1);

        XWPFTemplate.compile(templatePath + "reference_picture_iterable.docx")
                .render(new HashMap<String, Object>() {
                    {
                        put("mores", mores);
                    }
                }).writeToFile(outPath + "out_reference_picture_iterable.docx");
    }

    @SuppressWarnings("serial")
    @Test
    public void testReplacePictureByOptionalText() throws Exception {
        XWPFTemplate.compile(templatePath + "reference_picture.docx")
                .render(new HashMap<String, Object>() {
                    {
                        put("img", Pictures.ofLocal(base + "sayi.png").size(100, 120).create());
                    }
                }).writeToFile(outPath + "out_reference_picture.docx");
    }

}
