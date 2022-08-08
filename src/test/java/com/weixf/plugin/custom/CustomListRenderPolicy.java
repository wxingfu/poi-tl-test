package com.weixf.plugin.custom;

import com.deepoove.poi.data.Numberings;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.render.WhereDelegate;

import java.util.List;

public class CustomListRenderPolicy extends AbstractRenderPolicy<List<String>> {

    @Override
    public void doRender(RenderContext<List<String>> context) throws Exception {
        // anywhere delegate
        WhereDelegate where = context.getWhereDelegate();
        // anything
        List<String> thing = context.getThing();
        // do 列表
        // where.renderNumbering(Numberings.of(thing.toArray(new String[] {})).create());
        String[] strings = thing.toArray(new String[]{});
        // do 列表
        where.renderNumbering(Numberings.ofDecimal(strings).create());
        // clear
        clearPlaceholder(context, true);
    }
}
