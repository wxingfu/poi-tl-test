package com.weixf.xwpf;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.FilePictureRenderData;
import com.deepoove.poi.data.HyperlinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

/**
 * @author Sayi
 */
public class TextboxTest {

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
        imgPath = base + "images/";
        markdownPath = base + "markdown/";
        outPath = base + "out/";
    }

    @Test
    public void testRender() throws Exception {
        Map<String, Object> datas = new HashMap<String, Object>() {
            {
                put("name", "Poi-tl");
                put("word", "模板引擎");
                put("time", "2019-05-31");
                put("author", new TextRenderData("000000", "Sayi卅一"));
                put("introduce", new HyperlinkTextRenderData("http://www.deepoove.com", "http://www.deepoove.com"));
                put("portrait", new FilePictureRenderData(60, 60, base + "sayi.png"));

            }
        };

        List<Map<String, Object>> mores = new ArrayList<Map<String, Object>>();
        mores.add(datas);
        mores.add(datas);
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("mores", mores);
//                put("volumes", "1900");
//                put("journal", new HashMap<String, String>(){
//                    {
//                        put("issn", "8848");
//                    }
//                });
            }
        };

        XWPFTemplate.compile(templatePath + "template_textbox.docx").render(data)
                .writeToFile(outPath + "out_template_textbox.docx");

    }

}
