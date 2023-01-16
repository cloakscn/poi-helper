package cn.cloaks.poihelper.handler;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.List;
import java.util.Map;

public class ParagraphHandler extends CursorHandler {

    private Map dataSource;

    public ParagraphHandler(Map dataSource) {
        this.dataSource = dataSource;
    }

    public void disposeParagraph(XWPFDocument xwpfDocument) {
        List<XWPFParagraph> paragraphs = xwpfDocument.getParagraphs();
        int size = paragraphs.size();
        for (int i = 0; i < size; i++) {
            XWPFParagraph paragraph = xwpfDocument.getParagraphArray(i);
            XWPFParagraph newParagraph;
            if (dispose(paragraph, dataSource)) {
                newParagraph = xwpfDocument.createParagraph();
                newParagraph.getCTP().setPPr(paragraph.getCTP().getPPr());
                XWPFRun run = newParagraph.createRun();
                run.setText(getContent());
            } else {
                newParagraph = xwpfDocument.createParagraph();
                copyParagraphStyle(newParagraph, paragraph);
            }
        }

        for (int i = size; i > 0; i--) {
            int posOfParagraph = xwpfDocument.getPosOfParagraph(paragraphs.get(i));
            xwpfDocument.removeBodyElement(posOfParagraph);
        }
    }
}
