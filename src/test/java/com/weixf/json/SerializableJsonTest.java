package com.weixf.json;


import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.deepoove.poi.data.*;
import com.deepoove.poi.data.style.Style;
import com.deepoove.poi.jsonmodel.support.DefaultGsonHandler;
import com.deepoove.poi.jsonmodel.support.GsonHandler;
import com.deepoove.poi.jsonmodel.support.GsonPreRenderDataCastor;
import com.google.gson.reflect.TypeToken;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Type;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertArrayEquals;
import static org.junit.jupiter.api.Assertions.assertEquals;

public class SerializableJsonTest {
    GsonHandler provider = new DefaultGsonHandler();
    String jsonStr = "";

    @Test
    void testTextRenderData() throws Exception {
        TextRenderData data = Texts.of("poi-tl").link("http://deepoove.com").color("00FF00").create();
        TextRenderData result = write(data).getResult(HyperlinkTextRenderData.class);

        assertEquals(result.getText(), data.getText());
        assertEquals(((HyperlinkTextRenderData) result).getUrl(), ((HyperlinkTextRenderData) data).getUrl());
        assertEquals(result.getStyle().getColor(), data.getStyle().getColor());
        assertEquals(result.getStyle().getUnderlinePatterns(), data.getStyle().getUnderlinePatterns());
    }

    @Test
    void testPictureRenderData() throws Exception {
        PictureRenderData data = Pictures.ofStream(new FileInputStream("src/test/resources/logo.png"))
                .size(100, 200)
                .altMeta("NO IMAGE")
                .create();
        PictureRenderData result = write(data).getResult(PictureRenderData.class);

        assertEquals(result.getPictureStyle().getWidth(), data.getPictureStyle().getWidth());
        assertEquals(result.getPictureStyle().getHeight(), data.getPictureStyle().getHeight());
        assertEquals(result.getPictureType(), data.getPictureType());
        assertEquals(result.getAltMeta(), data.getAltMeta());
    }

    @Test
    void testChartRenderData() throws Exception {
        ChartMultiSeriesRenderData data = Charts
                .ofComboSeries("ComboChartTitle", new String[]{"??????", "English", "?????????", "portugu??s"})
                .addBarSeries("countries", new Double[]{15.0, 6.0, 18.0, 231.0})
                .create();

        ChartMultiSeriesRenderData result = write(data).getResult(ChartMultiSeriesRenderData.class);

        assertEquals(result.getChartTitle(), data.getChartTitle());
        assertArrayEquals(result.getCategories(), data.getCategories());
        assertEquals(result.getSeriesDatas().get(0).getName(), data.getSeriesDatas().get(0).getName());
        assertEquals(result.getSeriesDatas().get(0).getComboType(), data.getSeriesDatas().get(0).getComboType());
    }

    @Test
    void testDocxRenderData() throws Exception {
        DocxRenderData data = new DocxRenderData(
                new File("src/test/resources/template.docx"));
        DocxRenderData result = write(data).getResult(DocxRenderData.class);

        assertArrayEquals(result.getMergedDoc(), data.getMergedDoc());
    }

    @Test
    void testTableRenderData() throws Exception {
        RowRenderData row = Rows.of(Cells.of("lisi").create(), Cells.of(Pictures.ofLocal("sayi.png").create()).create())
                .create();
        TableRenderData data = Tables.of(row).width(10.01f, null).center().create();

        TableRenderData result = write(data).getResult(TableRenderData.class);

        assertEquals(result.getTableStyle().getWidth(), data.getTableStyle().getWidth());
        assertEquals(result.getTableStyle().getAlign(), data.getTableStyle().getAlign());
        checkParagraph(result.getRows().get(0).getCells().get(0).getParagraphs().get(0),
                data.getRows().get(0).getCells().get(0).getParagraphs().get(0));

    }

    @Test
    void testParagraphRenderData() throws Exception {
        ParagraphRenderData data = Paragraphs.of("boshi").addPicture(Pictures.ofLocal("sayi.png").create()).create();

        ParagraphRenderData result = write(data).getResult(ParagraphRenderData.class);
        checkParagraph(result, data);

    }

    private void checkParagraph(ParagraphRenderData result, ParagraphRenderData data) {
        List<RenderData> contents = data.getContents();
        List<RenderData> contents2 = result.getContents();
        assertEquals(contents2.size(), contents.size());
        for (int i = 0; i < contents2.size(); i++) {
            RenderData renderData2 = contents2.get(i);
            RenderData renderData = contents.get(i);
            assertEquals(renderData2.getClass(), renderData.getClass());
            if (renderData instanceof TextRenderData) {
                assertEquals(((TextRenderData) renderData2).getText(), ((TextRenderData) renderData).getText());
            }
        }

    }

    @Test
    void testNumberingRenderData() throws Exception {
        Style fmtStyle = Style.builder().buildColor("00FF00").build();
        NumberingRenderData data = Numberings.ofBullet()
                .addItem(new TextRenderData("df2d4f", "Deeply in love with the things you love, just deepoove."))
                .addItem(Paragraphs.of()
                        .addText("Deeply in love with the things you love, just deepoove.")
                        .glyphStyle(fmtStyle)
                        .create())
                .addItem(new TextRenderData("5285c5", "Deeply in love with the things you love, just deepoove."))
                .create();

        NumberingRenderData result = write(data).getResult(NumberingRenderData.class);

        assertArrayEquals(result.getFormats().toArray(), data.getFormats().toArray());
        for (int i = 0; i < 3; i++) {
            TextRenderData dataTxt = (TextRenderData) data.getItems().get(i).getItem().getContents().get(0);
            TextRenderData resultTxt = (TextRenderData) data.getItems().get(i).getItem().getContents().get(0);
            assertEquals(resultTxt.getText(), dataTxt.getText());
        }
        assertEquals(result.getItems().get(1).getItem().getParagraphStyle().getGlyphStyle().getColor(),
                data.getItems().get(1).getItem().getParagraphStyle().getGlyphStyle().getColor());
    }

    @Test
    void testAll() throws Exception {
        // create table
        RowRenderData header = Rows.of("Word????????????", "???????????????", "?????????")
                .textColor("FFFFFF")
                .bgColor("ff9800")
                .center()
                .rowHeight(2.5f)
                .create();
        RowRenderData row0 = Rows.create("Poi-tl", "???Java??????????????????", "????????????????????????????????????POI?????????????????????");
        RowRenderData row1 = Rows.create("Apache Poi", "???Java??????????????????", "????????????????????????????????????");
        RowRenderData row2 = Rows.create("Freemarker", "XML??????????????????", "?????????????????????XML??????");
        RowRenderData row3 = Rows.create("OpenOffice", "????????????OpenOffice??????", "?????????????????????OpenOffice???API");
        TableRenderData table = Tables.create(header, row0, row1, row2, row3);

        Map<String, Object> datas = new HashMap<String, Object>();
        // text
        datas.put("header", "Deeply love what you love.");
        datas.put("name", "Poi-tl");
        datas.put("word", "????????????");
        datas.put("time", "2020-08-31");
        datas.put("what", "Java Word??????????????? Minimal Microsoft word(docx) templating with {{template}} in Java.");
        datas.put("author", Texts.of("Sayi??????").color("000000").create());

        // hyperlink
        datas.put("introduce", Texts.of("http://www.deepoove.com").link("http://www.deepoove.com").create());
        // picture
        datas.put("portrait", Pictures.ofLocal("src/test/resources/sayi.png").size(60, 60).create());
        // table
        datas.put("solution_compare", table);
        // numbering
        datas.put("feature",
                Numberings.create("Plug-in grammar, add new grammar by yourself",
                        "Supports word text, local pictures, web pictures, table, list, header, footer...",
                        "Templates, not just templates, but also style templates"));

        // chart
        datas.put("chart",
                Charts.ofMultiSeries("?????????", new String[]{"?????????", "?????????"})
                        .addSeries("poi-tl", new Double[]{10.0, 5.0})
                        .addSeries("freemark", new Double[]{90.0, 70.0})
                        .create());

        Map<String, Object> result = write(datas).getResult(new TypeToken<Map<String, Object>>() {
        }.getType());

        Configure config = Configure.builder().addPreRenderDataCastor(new GsonPreRenderDataCastor()).build();

        XWPFTemplate.compile("src/test/resources/template.docx", config)
                .render(result)
                .writeToFile("target/out_template_serializable.docx");

    }

    private SerializableJsonTest write(Object data) throws IOException {
        jsonStr = provider.writer().toJson(data);
//        System.out.println(jsonStr);
        return this;
    }

    private <T> T getResult(Class<T> clazz) throws IOException, ClassNotFoundException {
        return provider.parser().fromJson(jsonStr, clazz);
    }

    private Map<String, Object> getResult(Type collectionType) {
        return provider.parser().fromJson(jsonStr, collectionType);
    }

}
