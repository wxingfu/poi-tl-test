package com.weixf.plugin.highlight;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.highlight.HighlightRenderData;
import com.deepoove.poi.plugin.highlight.HighlightRenderPolicy;
import com.deepoove.poi.plugin.highlight.HighlightStyle;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Properties;

public class HighlightRenderPolicyTest {


    public static String base;
    public static String templatePath;
    public static String imgPath;
    public static String markdownPath;
    public static String hightlightPath;
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
        hightlightPath = base + "hightlight/";
        outPath = base + "out/";
    }

    @Test
    public void testHightlight() throws IOException {
        HighlightRenderData code = new HighlightRenderData();
        code.setCode("/**\n"
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
        code.setLanguage("java");
        code.setStyle(HighlightStyle.builder().withShowLine(true).build());

        List<Map<String, Object>> codes = new ArrayList<>();
        String[] themes = new String[]{"github", "idea", "zenburn", "far", "androidstudio", "solarized-light",
                "solarized-dark", "xcode", "vs", "tomorrow", "agate", "arduino-light", "darcula", "dark", "default",
                "docco", "dracula", "foundation", "googlecode", "monokai", "monokai-sublime", "mono-blue", "ocean",
                "rainbow", "stackoverflow-dark", "stackoverflow-light", "tomorrow-night", "a11y-dark", "a11y-light",
                "an-old-hope", "arta", "ascetic", "atom-one-dark", "atom-one-light", "gml", "gradient-dark",
                "magula", "nord"};
        for (String theme : themes) {
            HighlightRenderData source = new HighlightRenderData();
            source.setCode("/**\n"
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
            source.setLanguage("java");
            source.setStyle(HighlightStyle.builder().withShowLine(true).withTheme(theme).build());

            codes.add(new HashMap<String, Object>() {
                {
                    put("theme", theme);
                    put("code", source);
                }
            });
        }

        Map<String, Object> data = new HashMap<>();
        data.put("code", code);
        data.put("themes", codes);

        Configure config = Configure.builder().bind("code", new HighlightRenderPolicy()).build();
        XWPFTemplate.compile(hightlightPath + "highlight_template.docx", config).render(data)
                .writeToFile(outPath + "out_highlight.docx");
    }

}
