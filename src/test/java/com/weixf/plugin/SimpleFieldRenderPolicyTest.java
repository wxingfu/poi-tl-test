package com.weixf.plugin;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.plugin.field.SimpleFieldRenderPolicy;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.HashMap;

public class SimpleFieldRenderPolicyTest {

    @Test
    public void testField() throws IOException {
        XWPFDocument xwpfDocument = new XWPFDocument();
        xwpfDocument.createParagraph().createRun().setText("{{field}}");

        HashMap<String, Object> data = new HashMap<String, Object>();
        data.put("field", "EQ \\o\\ac(□, 1)");
        Configure config = Configure.builder().bind("field", new SimpleFieldRenderPolicy()).build();

        XWPFTemplate template = XWPFTemplate.compile(xwpfDocument, config).render(data);
        template.writeToFile("target/out_render_field.docx");
    }

}
