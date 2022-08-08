package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.TextRenderData;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class JSONRenderPolicyTest {

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
    public void testJSONRender() throws Exception {

        File file = new File(base + "petstore.json");
        FileInputStream in = new FileInputStream(file);
        int size = in.available();
        byte[] buffer = new byte[size];
        in.read(buffer);
        in.close();

        ObjectMapper objectMapper = new ObjectMapper();
        JsonNode jsonNode = objectMapper.readTree(new String(buffer));
        List<TextRenderData> codes = new JSONRenderPolicy("000000").convert(jsonNode, 1);
        codes.forEach(System.out::print);

        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("json", new String(buffer));
                put("codes", codes);
            }
        };

        JSONRenderPolicy policy = new JSONRenderPolicy("ffffff");
        Configure config = Configure.builder().bind("json", policy).build();
        XWPFTemplate template =
                XWPFTemplate.compile(templatePath + "render_json.docx", config)
                        .render(data);

        template.writeToFile(outPath + "out_render_json.docx");

    }
}
