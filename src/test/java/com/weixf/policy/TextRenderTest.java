package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.Texts;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.xwpf.XWPFHighlightColor;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Text Render test case")
public class TextRenderTest {

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
    public void testTextRender() throws Exception {

        Map<String, Object> data = new HashMap<String, Object>();
        data.put("title", "Hello, poi tl.");
        data.put("text", new TextRenderData("28a745", "我是\n绿色且换行的文字"));
        // 超链接
        data.put("link", Texts.of("Deepoove website.").link("http://www.deepoove.com").bold().create());
        // 邮箱超链接
        data.put("maillink", Texts.of("发邮件给作者").mailto("sayi@apache.org", "poi-tl").create());
        // 指定文本样式
        Style style = new Style("FF5722");
        style.setBold(true);
        style.setFontSize(48L);
        style.setItalic(true);
        style.setStrike(true);
        style.setUnderlinePatterns(UnderlinePatterns.SINGLE);
        style.setFontFamily("微软雅黑");
        style.setWesternFontFamily("Monaco");
        style.setCharacterSpacing(20);
        style.setHighlightColor(XWPFHighlightColor.DARK_GREEN);
        data.put("word", Texts.of("深爱你所爱；just deepoove.").style(style).create());
        // 换行
        data.put("newline", "hi\n\n\n\n\nhello");
        // 从文件读取文字
        File file = new File(templatePath + "render_text_word.txt");
        FileInputStream in = new FileInputStream(file);
        int size = in.available();
        byte[] buffer = new byte[size];
        in.read(buffer);
        in.close();
        data.put("newline_from_file", new String(buffer));
        // 上标
        data.put("super", Texts.of("a+b").sup().create());
        // 下标
        data.put("sub", Texts.of("a+b").sub().create());

        XWPFTemplate.compile(templatePath + "render_text.docx")
                .render(data)
                .writeToFile(outPath + "out_render_text.docx");

    }

}
