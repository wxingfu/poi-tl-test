package com.weixf.policy;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.MergeCellRule.Grid;
import com.deepoove.poi.policy.TableRenderPolicy;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Table Render test case")
public class TableRenderTest {

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
    public void testTableByBuilder() throws Exception {
        // table
        RowRenderData row0, row1, row2;
        CellRenderData cell, cell1, cell2;

        cell = Cells.of(Texts.of("Hello").color("ffffff").create())
                .addParagraph(Paragraphs.of(Texts.of("World").color("ffffff").bold().create()).left().create())
                .create();
        cell1 = Cells.create("这是单元格数据");
        cell2 = Cells.of(Pictures.ofLocal(base + "sayi.png").size(40, 40).create()).addParagraph("Sayi")
                .center().create();

        row0 = Rows.of(cell, cell, cell, cell).center().rowHeight(2.54f).bgColor("f58d71").create();
        row1 = Rows.create(cell2, cell1, cell1, cell1);
        row2 = Rows.create(cell2, cell1, cell1, cell1);

        TableRenderData table = Tables.of(row0, row1, row2).width(14.63f, new double[]{5.63f, 3.0f, 3.0f, 3.0f})
                .create();

        // table 1
        RowRenderData rowNoData = Rows.create("没有数据", null, null, null);
        RowRenderData header = Rows.of("列0", "列1", "列2", "列3").horizontalCenter().bgColor("f58d71").create();
        TableRenderData table1 = Tables.of(header, rowNoData).create();
        MergeCellRule rule = MergeCellRule.builder().map(Grid.of(1, 0), Grid.of(1, 3)).build();
        table1.setMergeRule(rule);

        // table 2
        TableRenderData table2 = Tables.of(new String[][]{new String[]{"00", "01", "02", "03", "04"},
                new String[]{"10", "11", "12", "13", "14"}, new String[]{"20", "21", "22", "23", "24"},
                new String[]{"30", "31", "32", "33", "34"}}).create();
        MergeCellRule rule2 = MergeCellRule.builder().map(Grid.of(0, 0), Grid.of(1, 1))
                .map(Grid.of(1, 2), Grid.of(2, 3)).map(Grid.of(2, 0), Grid.of(2, 1)).map(Grid.of(1, 4), Grid.of(3, 4))
                .map(Grid.of(3, 1), Grid.of(3, 2)).build();
        table2.setMergeRule(rule2);

        Map<String, Object> data = new HashMap<String, Object>();
        data.put("table", table);
        data.put("no_data_table", table1);
        data.put("complex_table", table2);

        // 不加标签类型 # ，依靠插件处理
        Configure config = Configure.builder()
                .bind("table", new TableRenderPolicy())
                .bind("complex_table", new TableRenderPolicy())
                .bind("no_data_table", new TableRenderPolicy())
                .build();
        XWPFTemplate.compile(templatePath + "render_tablev1.docx", config)
                .render(data)
                .writeToFile(outPath + "out_render_tablev1.docx");

        // 加标签类型 #
        XWPFTemplate.compile(templatePath + "render_tablev2.docx")
                .render(data)
                .writeToFile(outPath + "out_render_tablev2.docx");

    }
}
