package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.FilePictureRenderData;
import com.deepoove.poi.data.Numberings;
import com.deepoove.poi.data.Tables;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.ListRenderPolicy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.*;

@DisplayName("List Render test case")
public class ListRenderPolicyTest {

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
    public void testListData() throws Exception {

        List<Object> list = new ArrayList<Object>();

        list.add(new TextRenderData("ver 0.0.3"));
        list.add(new FilePictureRenderData(100, 120, base + "logo.png"));
        list.add(new TextRenderData("9d55b8", "Deeply in love with the things you love, just deepoove."));
        list.add(new TextRenderData("ver 0.0.4"));
        list.add(new FilePictureRenderData(100, 120, base + "logo.png"));
        list.add(Numberings.ofDecimalParentheses()
                .addItem("Deeply in love with the things you love, just deepoove.")
                .addItem("Deeply in love with the things you love, just deepoove.")
                .addItem("Deeply in love with the things you love, just deepoove.").create());
        list.add(Tables.of(new String[][]{
                new String[]{"00", "01"},
                new String[]{"10", "11"},}).create());

        Map<String, Object> datas = new HashMap<String, Object>();
        datas.put("website", list);


        // List demo {{website}} List demo.
        Configure config = Configure.builder().bind("website", new ListRenderPolicy() {
        }).build();
        XWPFTemplate.compile(templatePath + "render_list.docx", config)
                .render(datas)
                .writeToFile(outPath + "out_render_list.docx");
    }

}
