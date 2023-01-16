package cn.cloaks.poihelper.handler;


import org.apache.poi.xwpf.usermodel.PositionInParagraph;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class CursorHandler {

    private String content;

    public boolean dispose(XWPFParagraph xwpfParagraph, Map dataSource) {
        try {
            List<XWPFRun> xwpfRuns = xwpfParagraph.getRuns();

            String regEx = "\\{.+?\\}";
            Pattern pattern = Pattern.compile(regEx);
            Matcher matcher = pattern.matcher(xwpfParagraph.getText());
            if (matcher.find()) {
                int beginRunIndex = xwpfParagraph.searchText("{", new PositionInParagraph()).getBeginRun();
                int endRunIndex = xwpfParagraph.searchText("}", new PositionInParagraph()).getEndRun();
                StringBuffer key = new StringBuffer();

                // 处理 ${**} 被分成多个run
                if (beginRunIndex == endRunIndex) {
                    // {**} 在一个run标签内
                    XWPFRun beginXwpfRun = xwpfRuns.get(beginRunIndex);
                    String beginXwpfRunText = beginXwpfRun.text();

                    int beginTagIndex = beginXwpfRunText.indexOf("{");
                    int endTagIndex = beginXwpfRunText.indexOf("}");
                    int length = beginXwpfRunText.length();

                    XWPFRun newRun = xwpfParagraph.insertNewRun(beginRunIndex);
                    newRun.getCTR().setRPr(beginXwpfRun.getCTR().getRPr());

                    if (beginTagIndex == 0 && endTagIndex == length - 1) {
                        // 该run标签只有 {**}
                        // 设置文本
                        key.append(beginXwpfRunText.substring(1, endTagIndex));
                        newRun.setText(getDataSource(key.toString(), dataSource).toString());
                    } else {
                        // 该run标签为**{**}** 或者 **{**} 或者{**}**，替换key后，还需要加上原始key前后的文本
                        // 设置文本
                        key.append(beginXwpfRunText.substring(beginXwpfRunText.indexOf("{") + 1, beginXwpfRunText.indexOf("}")));
                        String textString = beginXwpfRunText.substring(0, beginTagIndex) + getDataSource(key.toString(), dataSource) + beginXwpfRunText.substring(endTagIndex + 1);
                        newRun.setText(textString);
                    }
                    xwpfParagraph.removeRun(beginRunIndex + 1);
                } else {
                    // {**} 被分成多个 run
                    // 先处理起始run标签,取得第一个{key}值
                    XWPFRun beginRun = xwpfRuns.get(beginRunIndex);
                    String beginRunText = beginRun.text();
                    int beginIndex = beginRunText.indexOf("{");

                    if (beginRunText.length() > 1) {
                        key.append(beginRunText.substring(beginIndex + 1));
                    }

                    // 需要移除的run
                    ArrayList<Integer> removeRunList = new ArrayList<>();
                    // 处理中间的run
                    for (int i = beginRunIndex + 1; i < endRunIndex; i++) {
                        XWPFRun run = xwpfRuns.get(i);
                        String runText = run.text();
                        key.append(runText);
                        removeRunList.add(i);
                    }
                    // 获取endRun中的key值
                    XWPFRun endRun = xwpfRuns.get(endRunIndex);
                    String endRunText = endRun.text();
                    int endIndex = endRunText.indexOf("}");

                    // run 中 **} 或者 **}**
                    if (endRunText.length() > 1 && endIndex != 0) {
                        key.append(endRunText.substring(0, endIndex));
                    }

                    // 取得 key 值后替换标签
                    // 先处理开始标签
                    if (beginRunText.length() == 2) {
                        // run标签内文本{
                        XWPFRun insertNewRun = xwpfParagraph.insertNewRun(beginRunIndex);
                        insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                        // 设置文本
                        insertNewRun.setText(getDataSource(key.toString(), dataSource).toString());
                        // 移除原始的run
                        xwpfParagraph.removeRun(beginRunIndex + 1);
                    } else {
                        // 该run标签为**{**或者 {** ，替换key后，还需要加上原始key前的文本
                        XWPFRun insertNewRun = xwpfParagraph.insertNewRun(beginRunIndex);
                        insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                        // 设置文本
                        String textString = beginRunText.substring(0, beginRunText.indexOf("{")) + getDataSource(key.toString(), dataSource);
                        insertNewRun.setText(textString);
                        xwpfParagraph.removeRun(beginRunIndex + 1);//移除原始的run
                    }

                    // 处理结束标签
                    if (endRunText.length() == 1) {
                        // run标签内文本只有}
                        XWPFRun insertNewRun = xwpfParagraph.insertNewRun(endRunIndex);
                        insertNewRun.getCTR().setRPr(endRun.getCTR().getRPr());
                        // 设置文本
                        insertNewRun.setText("");
                        xwpfParagraph.removeRun(endRunIndex + 1);//移除原始的run

                    } else {
                        // 该run标签为**}**或者 }** 或者**}，替换key后，还需要加上原始key后的文本
                        XWPFRun insertNewRun = xwpfParagraph.insertNewRun(endRunIndex);
                        insertNewRun.getCTR().setRPr(endRun.getCTR().getRPr());
                        // 设置文本
                        String textString = endRunText.substring(endRunText.indexOf("}") + 1);
                        insertNewRun.setText(textString);
                        xwpfParagraph.removeRun(endRunIndex + 1);//移除原始的run
                    }

                    //处理中间的run标签
                    for (int i = 0; i < removeRunList.size(); i++) {
                        XWPFRun xWPFRun = xwpfRuns.get(removeRunList.get(i));//原始run
                        XWPFRun insertNewRun = xwpfParagraph.insertNewRun(removeRunList.get(i));
                        insertNewRun.getCTR().setRPr(xWPFRun.getCTR().getRPr());
                        insertNewRun.setText("");
                        xwpfParagraph.removeRun(removeRunList.get(i) + 1);//移除原始的run
                    }
                }
                content = xwpfParagraph.getText();
                return true;
            }
            return false;
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }


    /**
     * 根据 key 获取数据源信息
     *
     * @param key        数据键
     * @param dataSource 数据源
     * @return
     */
    private Object getDataSource(String key, Map<String, Object> dataSource) {
        Object result = "";

        if (key.isEmpty()) {
            return result;
        }

        result = parseKey(key.split("\\."), dataSource);
        if (result == null) {
//            log.error("数据源缺失：" + key);
        }

        return result;
    }

    private Object parseKey(String[] split, Map<String, Object> dataSource) {
        if (split.length == 0 || split[0].isEmpty()) {
            return null;
        }

        Object o = dataSource.get(split[0]);

        if (o == null) {
            return null;
        }

        if (split.length == 1) {
            return o;
        } else {
            return parseKey(Arrays.copyOfRange(split, 1, split.length), (Map<String, Object>) o);
        }
    }

    /**
     * 复制文本段落 XWPFParagraph 格式
     *
     * @param desXwpfParagraph 目标模板段落段落
     * @param srcXwpfParagraph 源模板段落节点
     */
    public void copyParagraphStyle(XWPFParagraph desXwpfParagraph, XWPFParagraph srcXwpfParagraph) {
        desXwpfParagraph.getCTP().setPPr(srcXwpfParagraph.getCTP().getPPr());
        for (int i = 0; i < desXwpfParagraph.getRuns().size(); i++) {
            desXwpfParagraph.removeRun(i);
        }
        for (XWPFRun srcRun : srcXwpfParagraph.getRuns()) {
            XWPFRun desRun = desXwpfParagraph.createRun();
            copyRunStyle(desRun, srcRun);
        }
    }

    /**
     * 复制文本节点 run
     *
     * @param desXwpfRun 目标模板文本节点
     * @param srcXwpfRun 源模板文本节点
     */
    private void copyRunStyle(XWPFRun desXwpfRun, XWPFRun srcXwpfRun) {
        desXwpfRun.getCTR().setRPr(srcXwpfRun.getCTR().getRPr());
        // 设置文本
        desXwpfRun.setText(srcXwpfRun.text());
    }

    public String getContent() {
        return content;
    }
}
