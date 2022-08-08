package com.weixf.plugin;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.HyperlinkTextRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.plugin.bookmark.BookmarkRenderPolicy;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;

@DisplayName("Bookmark Render test case")
public class BookmarkRenderPolicyTest {

    @Test
    public void testBookmarkRender() throws Exception {

        Map<String, Object> datas = new HashMap<String, Object>();
        datas.put("anchor1", new HyperlinkTextRenderData("poi-tl", "anchor:poi-tl"));
        datas.put("anchor2", new HyperlinkTextRenderData("文字", "anchor:我是绿色且换行的文字"));
        datas.put("title", "poi-tl");
        datas.put("text", new TextRenderData("28a745", "我是绿色且换行的文字"));

        BookmarkRenderPolicy bookmarkRenderPolicy = new BookmarkRenderPolicy();
        Configure config = Configure.builder().bind("title", bookmarkRenderPolicy).bind("text", bookmarkRenderPolicy)
                .build();
        XWPFTemplate.compile("src/test/resources/template/render_bookmark.docx", config).render(datas)
                .writeToFile("target/out_render_bookmark.docx");

    }

}
