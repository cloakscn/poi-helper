package cn.cloaks.poihelper.util;

import cn.cloaks.poihelper.handler.*;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.Map;

public class XWPFUtil {

    private XWPFDocument xwpfDocument;

    private ParagraphHandler paragraphHandler;

    private TableHandler tablehandler;

    private ChartHandler chartHandler;

    private HeaderHandler headerHandler;

    private FooterHandler footerHandler;

    private Map dataSource;

    private void init(InputStream is) {
        try {
            xwpfDocument = new XWPFDocument(is);
            paragraphHandler = new ParagraphHandler(dataSource);
            tablehandler = new TableHandler();
            chartHandler = new ChartHandler();
            headerHandler = new HeaderHandler();
            footerHandler = new FooterHandler();
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }

    public XWPFUtil(InputStream is, Map dataSource) {
        this.dataSource = dataSource;
        init(is);
    }

    public void replace() {
        paragraphHandler.disposeParagraph(xwpfDocument);
        tablehandler.dispose(xwpfDocument);
        chartHandler.dispose(xwpfDocument);
        headerHandler.dispose(xwpfDocument);
        footerHandler.dispose(xwpfDocument);
    }

    public void write(OutputStream os) {
        try {
            xwpfDocument.write(os);
        } catch (IOException e) {
            throw new RuntimeException(e);
        }
    }
}
