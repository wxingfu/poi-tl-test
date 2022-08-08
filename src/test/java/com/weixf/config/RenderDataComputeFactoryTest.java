package com.weixf.config;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.render.compute.RenderDataComputeFactory;
import com.google.common.collect.Maps;
import com.weixf.source.XWPFTestSupport;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.InputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class RenderDataComputeFactoryTest {

    @Test
    public void testConfigRenderDataCompute() throws IOException {
        RenderDataComputeFactory renderDataComputeFactory = model -> el -> "123";
        Configure config = Configure.builder().setRenderDataComputeFactory(renderDataComputeFactory).build();

        XWPFDocument doc = new XWPFDocument();
        doc.createParagraph().createRun().setText("{{name}}");
        InputStream inputStream = XWPFTestSupport.readInputStream(doc);

        XWPFTemplate template = XWPFTemplate.compile(inputStream, config).render(Maps.newHashMap());
        XWPFDocument document = XWPFTestSupport.readNewDocument(template);
        assertEquals("123", document.getParagraphArray(0).getText());

    }

}
