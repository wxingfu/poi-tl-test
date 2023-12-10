package com.weixf.resolver;

import com.deepoove.poi.config.Configure;
import com.deepoove.poi.resolver.DefaultElementTemplateFactory;
import com.deepoove.poi.resolver.TemplateResolver;
import com.deepoove.poi.template.run.RunTemplate;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.regex.Pattern;

import static org.junit.jupiter.api.Assertions.assertEquals;

@DisplayName("Create RunTemplate test case")
public class DefaultRunTemplateFactoryTest {

    @Test
    public void testCreateRunTemplate() {
        Configure config = Configure.builder().build();
        TemplateResolver resolver = new TemplateResolver(config);

        Pattern templatePattern = resolver.getTemplatePattern();
        Pattern gramerPattern = resolver.getGramerPattern();

        DefaultElementTemplateFactory runTemplateFactory = new DefaultElementTemplateFactory();

        String tag = "";
        RunTemplate template = null;

        String text = "{{/}}";
        if (templatePattern.matcher(text).matches()) {
            tag = gramerPattern.matcher(text).replaceAll("").trim();
            template = runTemplateFactory.createRunTemplate(config, tag, null);
        }
        assertEquals(tag, "/");
        assertEquals(template.toString(), text);
        assertEquals(template.getTagName(), "");

        text = "{{}}";
        if (templatePattern.matcher(text).matches()) {
            tag = gramerPattern.matcher(text).replaceAll("").trim();
            template = runTemplateFactory.createRunTemplate(config, tag, null);
        }
        assertEquals(tag, "");
        assertEquals(template.toString(), text);
        assertEquals(template.getTagName(), "");

        text = "{{name}}";
        if (templatePattern.matcher(text).matches()) {
            tag = gramerPattern.matcher(text).replaceAll("").trim();
            template = runTemplateFactory.createRunTemplate(config, tag, null);
        }
        assertEquals(tag, "name");
        assertEquals(template.toString(), text);
        assertEquals(template.getTagName(), "name");

        text = "{{?name}}";
        if (templatePattern.matcher(text).matches()) {
            tag = gramerPattern.matcher(text).replaceAll("").trim();
            template = runTemplateFactory.createRunTemplate(config, tag, null);
        }
        assertEquals(tag, "?name");
        assertEquals(template.toString(), text);
        assertEquals(template.getTagName(), "name");
    }

}
