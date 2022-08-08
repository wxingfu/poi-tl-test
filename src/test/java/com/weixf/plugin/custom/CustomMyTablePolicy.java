package com.weixf.plugin.custom;

import com.deepoove.poi.data.style.TableStyle;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import com.deepoove.poi.util.TableTools;
import com.deepoove.poi.xwpf.BodyContainer;
import com.deepoove.poi.xwpf.BodyContainerFactory;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.poi.xwpf.usermodel.XWPFTable;

import java.util.List;

public class CustomMyTablePolicy extends AbstractRenderPolicy<List<String>> {
    @Override
    public void doRender(RenderContext<List<String>> context) {
        // anywhere
        XWPFRun where = context.getWhere();
        // anything
        List<String> thing = context.getThing();
        // do 表格
        int row = thing.size() + 1, col = 2;
        BodyContainer bodyContainer = BodyContainerFactory.getBodyContainer(where);
        XWPFTable table = bodyContainer.insertNewTable(where, row, col);
        TableTools.widthTable(table, 14.65f, col);
        TableTools.borderTable(table, 4);
        table.getRow(0).getCell(0).setText("编号");
        table.getRow(0).getCell(1).setText("地点信息");
        for (int i = 0; i < thing.size(); i++) {
            table.getRow(i + 1).getCell(0).setText("No." + i);
            table.getRow(i + 1).getCell(1).setText(thing.get(i));
        }
        // clear
        clearPlaceholder(context, true);
    }
}
