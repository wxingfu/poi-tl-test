package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.Cells;
import com.deepoove.poi.data.MergeCellRule;
import com.deepoove.poi.data.MergeCellRule.Grid;
import com.deepoove.poi.data.RowRenderData;
import com.deepoove.poi.data.Rows;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.Tables;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.data.Texts;
import com.weixf.source.CustomTableRenderPolicy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileOutputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Table Render test case")
public class MiniTableRenderTest {

    public static String base;
    public static String templatePath;
    public static String imgPath;
    public static String markdownPath;
    public static String outPath;

    RowRenderData header, row0, row1, row2, row3;

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

        header = Rows.of(new TextRenderData("FFFFFF", "姓名"), new TextRenderData("FFFFFF", "学历")).bgColor("009688")
                .center().create();
        row0 = Rows.of(Texts.of("张三").link("http://deepoove.com").create(), new TextRenderData("1E915D", "研究生"))
                .create();
        row1 = Rows.create("李四", "博士");
        row2 = Rows.create("王五", "博士后");
        row3 = Rows.of(Cells.of(new TextRenderData("FFFFFF", "白字\n蓝底居左")).bgColor("0000ff").create(),
                        Cells.of(new TextRenderData("00ff00", "绿字灰底居右")).horizontalRight().bgColor("666666").create())
                .create();
    }

    @Test
    public void testTable() throws Exception {

        Map<String, Object> data = new HashMap<String, Object>();

        // 有表格头 有数据，宽度自适应
        data.put("table", Tables.create(header, row0, row1, row2, row3));
        // 没有表格头 没有数据，最终不会渲染这个table
        data.put("no_table", Tables.ofPercentWidth("100%").create());
        // 有数据，没有表格头
        data.put("no_header_table", Tables.create(row0, row1, row2, row3));
        // 有表格头 没有数据
        RowRenderData rowNoData = Rows.create("没有数据", null);
        TableRenderData table1 = Tables.of(header, rowNoData).create();
        MergeCellRule rule = MergeCellRule.builder().map(Grid.of(1, 0), Grid.of(1, 1)).build();
        table1.setMergeRule(rule);
        data.put("no_content_table", table1);

        // 指定宽度的表格, 表格居中
        TableRenderData miniTableRenderData =
                Tables.of(header, row0, row1, row2, row3)
                        .width(8.00f, null)
                        .center()
                        .create();
        data.put("width_table", miniTableRenderData);

        // 自定义表格策略
        // {{report}}
        Configure config = Configure.builder().bind("report", new CustomTableRenderPolicy()).build();
        XWPFTemplate template =
                XWPFTemplate.compile(templatePath + "render_table.docx", config)
                        .render(data);

        FileOutputStream out = new FileOutputStream(outPath + "out_render_table.docx");
        template.write(out);
        out.flush();
        out.close();
        template.close();
    }
}
