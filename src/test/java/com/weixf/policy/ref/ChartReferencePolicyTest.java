package com.weixf.policy.ref;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.data.ChartMultiSeriesRenderData;
import com.deepoove.poi.data.ChartSingleSeriesRenderData;
import com.deepoove.poi.data.Charts;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

@DisplayName("Chart test case")
public class ChartReferencePolicyTest {

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
    public void testChart() throws Exception {
        Map<String, Object> data = new HashMap<String, Object>();
        ChartMultiSeriesRenderData chart = createMultiSeriesChart();
        data.put("barChart", chart);
        data.put("VBarChart", chart);
        data.put("3dBarChart", chart);
        data.put("lineChart", chart);
        data.put("redarChart", chart);
        data.put("areaChart", chart);

        ChartMultiSeriesRenderData combChart = createComboChart();
        data.put("combChart", combChart);

        ChartSingleSeriesRenderData pie = createSingleSeriesChart();
        data.put("pieChart", pie);
        data.put("doughnutChart", pie);

        ChartMultiSeriesRenderData scatter = createMultiSeriesScatterChart();
        data.put("scatterChart", scatter);
        data.put("relatechart", scatter);

        XWPFTemplate template = XWPFTemplate.compile(templatePath + "reference_chart.docx").render(data);
        template.writeToFile(outPath + "out_reference_chart.docx");
    }

    private ChartSingleSeriesRenderData createSingleSeriesChart() {
        return Charts.ofSingleSeries("ChartTitle", new String[]{"俄罗斯", "加拿大", "美国", "中国", "巴西", "澳大利亚", "印度"})
                .series("countries", new Integer[]{17098242, 9984670, 9826675, 9596961, 8514877, 7741220, 3287263})
                .create();
    }

    private ChartMultiSeriesRenderData createMultiSeriesScatterChart() {
        return Charts.ofMultiSeries("ChartTitle", new String[]{"1", "3", "4", "6", "9", "2", "4"})
                .addSeries("A", new Integer[]{12, 4, 9, 2, 10, 5, 7})
                .addSeries("B", new Integer[]{2, 6, 3, 6, 2, 6, 9})
                .setxAsixTitle("X Title")
                .setyAsixTitle("Y Title")
                .create();
    }

    private ChartMultiSeriesRenderData createMultiSeriesChart() {
        return Charts
                .ofMultiSeries("CustomTitle", new String[]{"中文", "English", "日本語", "português", "中文", "English", "日本語", "português"})
                .addSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0, 150.0, 6.0, 118.0, 31.0})
                .addSeries("speakers", new Double[]{223.0, 119.0, 154.0, 142.0, 223.0, 119.0, 54.0, 42.0})
                .addSeries("youngs", new Double[]{323.0, 89.0, 54.0, 42.0, 223.0, 119.0, 54.0, 442.0})
                .create();
    }

    private ChartMultiSeriesRenderData createComboChart() {
        return Charts.ofComboSeries("ComboChartTitle", new String[]{"中文", "English", "日本語", "português"})
                .addBarSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0})
                .addBarSeries("speakers", new Double[]{223.0, 119.0, 154.0, 142.0})
                .addBarSeries("NewBar", new Double[]{223.0, 119.0, 154.0, 142.0})
                .addLineSeries("youngs", new Double[]{323.0, 89.0, 54.0, 42.0})
                .addLineSeries("NewLine", new Double[]{123.0, 59.0, 154.0, 42.0})
                .create();
    }

}
