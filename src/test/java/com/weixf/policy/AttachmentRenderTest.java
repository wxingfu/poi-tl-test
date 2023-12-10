package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.AttachmentType;
import com.deepoove.poi.data.Attachments;
import com.deepoove.poi.data.Charts;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.Texts;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.policy.AttachmentRenderPolicy;
import com.deepoove.poi.xwpf.XWPFHighlightColor;
import org.apache.poi.xwpf.usermodel.UnderlinePatterns;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.util.Arrays;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Attachment Render test case")
public class AttachmentRenderTest {

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
    public void testAttachmentRender() throws Exception {

        Map<String, Object> attachDatas = new HashMap<String, Object>() {
            {
                put("title", "Hello, poi tl.");
                put("text", new TextRenderData("28a745", "我是\n绿色且换行的文字"));
                // 超链接
                put("link", Texts.of("Deepoove website.").link("http://www.deepoove.com").bold().create());
                // 邮箱超链接
                put("maillink", Texts.of("发邮件给作者").mailto("sayi@apache.org", "poi-tl").create());
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
                put("word", Texts.of("深爱你所爱；just deepoove.").style(style).create());
                // 换行
                put("newline", "hi\n\n\n\n\nhello");
                // 从文件读取文字
                File file = new File(templatePath + "render_text_word.txt");
                FileInputStream in = new FileInputStream(file);
                int size = in.available();
                byte[] buffer = new byte[size];
                in.read(buffer);
                in.close();
                put("newline_from_file", new String(buffer));
                // 上标
                put("super", Texts.of("a+b").sup().create());
                // 下标
                put("sub", Texts.of("a+b").sub().create());
            }
        };

        XWPFTemplate attachTemplate =
                XWPFTemplate.compile(templatePath + "render_text.docx")
                        .render(attachDatas);

        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("attachment", Attachments.ofLocal(templatePath + "attachment.docx", AttachmentType.DOCX).create());
                put("xlsx", Attachments.ofLocal(templatePath + "attachment.xlsx", AttachmentType.XLSX).create());
                put("chart", Charts.ofMultiSeries("CustomTitle",
                                new String[]{"中文", "English", "日本語", "português", "中文", "English", "日本語", "português"})
                        .addSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0, 150.0, 6.0, 118.0, 31.0})
                        .addSeries("speakers", new Double[]{223.0, 119.0, 154.0, 142.0, 223.0, 119.0, 54.0, 42.0})
                        .addSeries("youngs", new Double[]{323.0, 89.0, 54.0, 42.0, 223.0, 119.0, 54.0, 442.0})
                        .create());
            }
        };

        Configure configure = Configure.builder()
                .bind("attachment", new AttachmentRenderPolicy())
                .bind("xlsx", new AttachmentRenderPolicy())
                .build();

        XWPFTemplate.compile(templatePath + "render_attachment.docx", configure)
                .render(new HashMap<String, Object>() {
                    {
                        put("attachment", Attachments.ofWordTemplate(attachTemplate).create());
                        put("xlsx", Attachments.ofLocal(templatePath + "attachment.xlsx", AttachmentType.XLSX).create());
                        put("list", Arrays.asList(data, data));
                    }
                }).writeToFile(outPath + "out_render_attachment.docx");

    }

}
