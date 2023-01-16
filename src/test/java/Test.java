import cn.cloaks.poihelper.util.XWPFUtil;
import org.apache.poi.xwpf.usermodel.Borders;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import org.apache.xmlbeans.XmlCursor;

import java.io.*;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;

public class Test {
    public static void main(String[] args) throws IOException {
        FileInputStream fileInputStream = new FileInputStream("C:\\workspace\\java\\poi-helper\\src\\test\\resources\\2023-01-14.docx");

        HashMap<String, Object> dataSource = new HashMap<>();
        dataSource.put("new_paragraph", "这是一个新的段落");
        XWPFUtil xwpfUtil = new XWPFUtil(fileInputStream, dataSource);
        xwpfUtil.replace();

        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("yyyy年MM月dd日 hh时mm分ss秒");
        String format = simpleDateFormat.format(new Date());

        FileOutputStream fileOutputStream = new FileOutputStream(String.format("C:\\workspace\\java\\poi-helper\\src\\test\\resources\\%s.docx", format));
        xwpfUtil.write(fileOutputStream);
        fileOutputStream.close();
    }
}
