package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.Pictures;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.Tables;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Width Fit Render test case")
public class TablePictureFitRenderTest {

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
    public void testFit() throws Exception {
        // table
        TableRenderData table = Tables
                .of(new String[][]{new String[]{"00", "01", "02", "03", "04"},
                        new String[]{"10", "11", "12", "13", "14"}})
                .percentWidth("100%", new int[]{10, 20, 10, 20, 40}).create();

        // picture
        PictureRenderData picture = Pictures.ofLocal(base + "large.png").fitSize().create();

        Map<String, Object> datas = new HashMap<String, Object>();
        datas.put("table", table);
        datas.put("picture", picture);

        Configure config = Configure.builder().build();
        XWPFTemplate.compile(templatePath + "width_fit.docx", config).render(datas)
                .writeToFile(outPath + "out_fit_width.docx");

    }
}
