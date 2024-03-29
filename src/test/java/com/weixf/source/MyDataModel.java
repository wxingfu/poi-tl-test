package com.weixf.source;

import com.deepoove.poi.data.NumberingRenderData;
import com.deepoove.poi.data.PictureRenderData;
import com.deepoove.poi.data.TableRenderData;
import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.expression.Name;

public class MyDataModel extends DataSourceBaseTest {

    private String word;
    private String time;
    private String what;
    private NumberingRenderData feature;
    @Name("solution_compare")
    private TableRenderData solutionCompare;
    private PictureRenderData portrait;
    private String author;
    private TextRenderData introduce;
    private String header;

    public String getWhat() {
        return this.what;
    }

    public void setWhat(String what) {
        this.what = what;
    }

    public NumberingRenderData getFeature() {
        return this.feature;
    }

    public void setFeature(NumberingRenderData feature) {
        this.feature = feature;
    }

    public TableRenderData getSolutionCompare() {
        return solutionCompare;
    }

    public void setSolutionCompare(TableRenderData solutionCompare) {
        this.solutionCompare = solutionCompare;
    }

    public PictureRenderData getPortrait() {
        return this.portrait;
    }

    public void setPortrait(PictureRenderData portrait) {
        this.portrait = portrait;
    }

    public String getAuthor() {
        return this.author;
    }

    public void setAuthor(String author) {
        this.author = author;
    }

    public TextRenderData getIntroduce() {
        return introduce;
    }

    public void setIntroduce(TextRenderData introduce) {
        this.introduce = introduce;
    }

    public String getHeader() {
        return this.header;
    }

    public void setHeader(String header) {
        this.header = header;
    }

    public String getWord() {
        return word;
    }

    public void setWord(String word) {
        this.word = word;
    }

    public String getTime() {
        return time;
    }

    public void setTime(String time) {
        this.time = time;
    }

}
