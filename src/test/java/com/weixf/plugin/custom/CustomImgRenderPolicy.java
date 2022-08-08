package com.weixf.plugin.custom;

import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.render.WhereDelegate;
import org.apache.poi.util.IOUtils;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;

public class CustomImgRenderPolicy extends AbstractRenderPolicy<String> {
    @Override
    public void doRender(RenderContext<String> context) throws Exception {
        // anywhere delegate
        WhereDelegate where = context.getWhereDelegate();
        // anything
        String thing = context.getThing();
        // do 图片
        FileInputStream stream = null;
        try {
            stream = new FileInputStream(thing);
            where.addPicture(stream, XWPFDocument.PICTURE_TYPE_PNG, 80, 80);
        } finally {
            IOUtils.closeQuietly(stream);
        }
        // clear
        clearPlaceholder(context, false);
    }
}
