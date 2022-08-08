package com.weixf.xwpf;

import com.deepoove.poi.util.NextIDUtils;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import org.apache.poi.xwpf.usermodel.*;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.math.BigInteger;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.Arrays;
import java.util.Calendar;
import java.util.TimeZone;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotEquals;

@DisplayName("Merge word test case")
public class NiceXWPFDocumentTest {

    @Test
    public void testMergeAtEnd() throws Exception {
        NiceXWPFDocument source = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_include.docx").toPath()));
        int sourceSize = source.getBodyElements().size();
        NiceXWPFDocument target = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_include_all.docx").toPath()));
        int targetSize = target.getBodyElements().size();
        source = source.merge(target);

        int resultSize = source.getBodyElements().size();
        // 会创建一个空的段落，然后移除
        assertEquals(sourceSize + targetSize, resultSize);

        source.close();
    }

    @Test
    public void testMergeAtRun() throws Exception {
        NiceXWPFDocument source = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_text.docx").toPath()));
        int sourceSize = source.getBodyElements().size();
        NiceXWPFDocument target1 = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_include_table.docx").toPath()));
        int targetSize1 = target1.getBodyElements().size();

        NiceXWPFDocument target2 = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_include_picture.docx").toPath()));
        int targetSize2 = target2.getBodyElements().size();
        source = source.merge(Arrays.asList(target1, target2), source.getParagraphArray(0).getRuns().get(0));

        int resultSize = source.getBodyElements().size();

        // 会移除source.getParagraphArray(0).getRuns().get(0)这个所在段落
        assertEquals(sourceSize - 1 + targetSize1 + targetSize2, resultSize);

        source.close();

    }

    @Test
    public void testMergeWithChart() throws Exception {
        NiceXWPFDocument source = new NiceXWPFDocument();
        NiceXWPFDocument target1 = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/reference_chart.docx").toPath()));

        source = source.merge(target1);
        source.write(Files.newOutputStream(Paths.get("target/out_merge_chart.docx")));
        source.close();

    }

    @Test
    public void testNewMergeNewNamespace() throws Exception {
        NiceXWPFDocument source;
        NiceXWPFDocument target;
        NiceXWPFDocument result;

        XWPFDocument doc = new XWPFDocument();
        ByteArrayOutputStream byteArrayOutputStream = new ByteArrayOutputStream();
        doc.write(byteArrayOutputStream);
        doc.close();

        source = new NiceXWPFDocument(new ByteArrayInputStream(byteArrayOutputStream.toByteArray()));
        assertEquals(-1, source.getXWPFDocument().getDocument().toString().indexOf(":wps="));

        target = new NiceXWPFDocument();
        assertNotEquals(-1, target.getXWPFDocument().getDocument().toString().indexOf(":wps="));
        XWPFParagraph createParagraph = target.createParagraph();
        createParagraph.createRun();

        result = source.merge(target);

        assertEquals(result.getParagraphs().size(), 1);
        assertEquals(result.getParagraphArray(0).getText(), "");
        assertNotEquals(-1, target.getXWPFDocument().getDocument().toString().indexOf(":wps="));

        source.close();
        target.close();
        result.close();

    }

    @Test
    public void testMergeSect() throws Exception {
        NiceXWPFDocument source = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_text.docx").toPath()));
        NiceXWPFDocument target = new NiceXWPFDocument(
                Files.newInputStream(new File("src/test/resources/template/render_include_sect.docx").toPath()));
        source = source.merge(target);

        source.write(Files.newOutputStream(Paths.get("target/out_merge_sect.docx")));
        source.close();
    }

    @Test
    public void testCreateComments() throws FileNotFoundException, IOException {
        NiceXWPFDocument document = new NiceXWPFDocument();
        XWPFComments docComments = document.createComments();
        XWPFComment addComment = docComments
                .createComment(NextIDUtils.getCommentMaxId(docComments).add(BigInteger.ONE));
        BigInteger cId = addComment.getCtComment().getId();

        XWPFParagraph paragraph = document.createParagraph();
        paragraph.getCTP().addNewCommentRangeStart().setId(cId); // comment range start is set before text run
        XWPFRun run = paragraph.createRun();
        run.setText("Paragraph with the first comment.");
        paragraph.getCTP().addNewCommentRangeEnd().setId(cId); // comment range end is set after text run

        paragraph.getCTP().addNewR().addNewCommentReference().setId(cId);

        addComment.setAuthor("Sayi");
        addComment.setInitials("s");
        addComment.setDate(Calendar.getInstance(TimeZone.getTimeZone("UTC")));
        addComment.createParagraph().createRun().setText("The first comment.");

        document.write(Files.newOutputStream(Paths.get("target/out_create_comments.docx")));
        document.close();
    }

}
