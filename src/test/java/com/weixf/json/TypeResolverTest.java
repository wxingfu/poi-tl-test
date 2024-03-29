package com.weixf.json;

import com.deepoove.poi.data.TextRenderData;
import com.deepoove.poi.policy.AbstractRenderPolicy;
import com.deepoove.poi.policy.TextRenderPolicy;
import com.deepoove.poi.render.RenderContext;
import net.jodah.typetools.TypeResolver;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertEquals;

public class TypeResolverTest {

    @Test
    public void testType() {
        Class<?>[] typeArguments = TypeResolver.resolveRawArguments(AbstractRenderPolicy.class, TextRenderPolicy.class);
        assertEquals(typeArguments[0], TextRenderData.class);

        typeArguments = TypeResolver.resolveRawArguments(AbstractRenderPolicy.class, MyRenderPolicy.class);
        assertEquals(typeArguments[0], TextRenderData.class);
    }

    abstract class MyAbstRenderPolicy<T> extends AbstractRenderPolicy<TextRenderData> {
    }

    class MyRenderPolicy extends MyAbstRenderPolicy<String> {

        @Override
        public void doRender(RenderContext<TextRenderData> context) throws Exception {
        }

    }

}
