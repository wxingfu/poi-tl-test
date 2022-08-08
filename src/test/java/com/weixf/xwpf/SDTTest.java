package com.weixf.xwpf;

import com.deepoove.poi.XWPFTemplate;
import org.apache.poi.ooxml.POIXMLProperties.CoreProperties;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.util.*;

/**
 * @author Sayi
 */
public class SDTTest {

//    paragraph.getIRuns().stream().filter(r -> r instanceof XWPFSDT).forEach(r -> {
//        ISDTContent isdtContent = ((XWPFSDT) r).getContent();
//        if (isdtContent instanceof XWPFSDTContent) {
//            @SuppressWarnings("unchecked")
//            List<ISDTContents> contents = (List<ISDTContents>) ReflectionUtils.getValue("bodyElements",
//                    isdtContent);
//            List<XWPFRun> collect = contents.stream()
//                    .filter(c -> c instanceof XWPFRun)
//                    .map(c -> (XWPFRun) c)
//                    .collect(Collectors.toList());
//            // to do refactor sdtcontent
//            resolveXWPFRuns(collect, metaTemplates, stack);
//        }
//    });

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
    public void testRenderSDTInParagraph() throws Exception {
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("titlefd", "Poi-tl");
                put("name", "模板引擎");
                put("list", new ArrayList<Map<String, Object>>() {
                    {
                        add(Collections.singletonMap("name", "Lucy"));
                        add(Collections.singletonMap("name", "Hanmeimei"));
                    }
                });
            }
        };
        XWPFTemplate.compile(templatePath + "template_sdt.docx")
                .render(data)
                .writeToFile(outPath + "out_sdt_para.docx");
    }

    @Test
    public void testRenderSDTBlockInBody() throws Exception {
        Map<String, Object> data = new HashMap<String, Object>() {
            {
                put("title", "Poi-tl");
                put("name", "模板引擎");
                put("list", new ArrayList<Map<String, Object>>() {
                    {
                        add(Collections.singletonMap("name", "Lucy"));
                        add(Collections.singletonMap("name", "Hanmeimei"));
                    }
                });
            }
        };
        XWPFTemplate.compile(templatePath + "sdt.docx")
                .render(data)
                .writeToFile(outPath + "out_sdt_block.docx");
    }

    @Test
    public void testRenderSDTInTextbox() throws Exception {
        Map<String, Object> data =
                new HashMap<String, Object>() {
                    {
                        put("A", "Poi-tl");
                    }
                };
        XWPFTemplate template = XWPFTemplate.compile(templatePath + "sdt_core.docx").render(data);
        CoreProperties coreProperties = template.getXWPFDocument().getProperties().getCoreProperties();
        coreProperties.setSubjectProperty("Poi-tl手册");
        template.writeToFile(outPath + "out_sdt_core.docx");
    }

    @Test
    public void testRenderSDTInTableRow() throws Exception {
        Map<String, Object> data =
                new HashMap<String, Object>() {
                    {
                        put("title", "Poi-tl");
                    }
                };
        XWPFTemplate template = XWPFTemplate.compile(templatePath + "sdt_cell.docx").render(data);
        CoreProperties coreProperties = template.getXWPFDocument().getProperties().getCoreProperties();
        coreProperties.setTitle("poi-tl");
        coreProperties.setDescription("desc");
        template.writeToFile(outPath + "out_sdt_cell.docx");
    }

}
