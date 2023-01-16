package cn.cloaks.poihelper.util;


import com.alibaba.fastjson2.JSONObject;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.Units;
import org.apache.poi.xddf.usermodel.chart.*;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.*;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.security.InvalidParameterException;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

/**
 * 版本声明: Copyright (c) 2022 FengTaiSEC Corporation.
 * 循环里边使用基于循环的路径标记变量
 *
 * @author 王艺 <wangyi@fengtaisec.com>
 * @brief 对docx文件中的文本及表格中的内容进行替换 --模板仅支持对 {key} 标签的替换
 * @date 2022/5/11 11:08
 * @history
 */
public class XWPFUtils {
    private static final String baseDataKey = "parametersMap";

    /**
     * 日志打印
     */
    private static final Logger log = LoggerFactory.getLogger(XWPFUtils.class);

    /**
     * 数据系列数据结构
     */
    public static class SeriesData {
        public SeriesData(String name, Series[] series) {
            this.name = name;
            this.series = series;
        }

        public static class Series {
            public Series(String name, Double[] values) {
                this.name = name;
                this.values = values;
            }

            /**
             * 系列名称
             */
            public String name;
            /**
             * 系列数据
             */
            public Double[] values;
        }

        /**
         * 系列名称
         */
        public String name;
        /**
         * 系列数据
         */
        public Series[] series;
    }


    /**
     * 分类数据结构
     */
    public static class CategoriesData {

        public CategoriesData(String name, String[] categories) {
            this.name = name;
            this.categories = categories;
        }

        /**
         * 分类名称
         */
        public String name;

        /**
         * 分类数据
         */
        public String[] categories;
    }

    /**
     * 柱状图数据结构
     */
    public static class ChartData {

        public ChartData() {
        }

        public ChartData(String title, ChartTypes chartTypes, CategoriesData categories, SeriesData seriesData) {
            this.title = title;
            this.chartType = chartTypes;
            this.categories = categories;
            this.seriesData = seriesData;
        }

        public String title;
        public ChartTypes chartType;
        public CategoriesData categories;
        public SeriesData seriesData;
    }

    /**
     * 文档属性
     */
    private XWPFDocument xwpfDocument;

    /**
     * 全局数据
     */
    private Map<String, Object> globalData, baseData;

    private int currentTableIndex, currentParagraphIndex, currentElementIndex;

    /**
     * 获取文档对象
     *
     * @return
     */
    public XWPFDocument getDocument() {
        return xwpfDocument;
    }

    /**
     * 设置文档对象
     */
    public void setDocument(XWPFDocument document) {
        this.xwpfDocument = document;
    }

    public XWPFUtils(InputStream inputStream, Map<String, Object> globalData) throws IOException {
        xwpfDocument = new XWPFDocument(inputStream);
        this.globalData = globalData;
        this.baseData = (Map<String, Object>) globalData.get(baseDataKey);
    }

    /**
     * 替换逻辑枚举类
     */
    private enum ReplaceTypes {
        IS_COPY,
        IS_LOOPS,
        IS_REPLACE,
        IS_NOT_LOOPS,
        TABLE_EXTERNAL_LOOPS,
        TABLE_INTERNAL_LOOPS,
    }

    /**
     * 把文档对象写入输出流中
     *
     * @throws IOException
     */
    public void write(OutputStream outputStream) throws IOException {
        xwpfDocument.write(outputStream);
    }

    private void disposeHeaders() {
        List<XWPFHeader> headerList = xwpfDocument.getHeaderList();
        for (XWPFHeader xwpfHeader : headerList) {
            List<XWPFParagraph> paragraphs = xwpfHeader.getParagraphs();
            for (XWPFParagraph xwpfParagraph : paragraphs) {
                replaceParagraph(xwpfParagraph, baseData);
            }
        }
    }

    private void disposeBodyElements() throws Exception {
        List<IBodyElement> bodyElementList = xwpfDocument.getBodyElements();
        // 当前操作表格对象的索引，当前操作段落对象的索引
        int bodyElementListSize = bodyElementList.size();


        // 处理文档元素
        while (currentElementIndex < bodyElementListSize) {
            IBodyElement bodyElement = bodyElementList.get(currentElementIndex);
            switch (bodyElement.getElementType()) {
                case TABLE:
                    // 处理表格
                    if (xwpfDocument.getTableArray(currentTableIndex).getText().contains("##{foreachParagraphsStart}##")) {
                        int[] arrays_index = disposeLoopsParagraphs(bodyElement, currentElementIndex, currentParagraphIndex, currentTableIndex, globalData);
                        currentElementIndex = arrays_index[0];
                        currentParagraphIndex = arrays_index[1];
                        currentTableIndex = arrays_index[2];
                    } else {
                        currentTableIndex = disposeTable(bodyElement, currentTableIndex, globalData, baseData);
                    }
                    break;
                case PARAGRAPH:
                    // 处理段落
                    currentParagraphIndex = disposeParagraph(bodyElement, currentParagraphIndex, baseData);
                    currentElementIndex++;
                    break;
                default:
            }
        }

        // 处理完毕模板，删除文本中的模板内容
        for (int i = 0; i < bodyElementListSize - 1; i++) {
            xwpfDocument.removeBodyElement(0);
        }
    }

    private int[] disposeLoopsParagraphs(IBodyElement bodyElement, int currentElementIndex, int currentParagraphIndex, int currentTableIndex, Map<String, Object> dataMap) throws Exception {
        XWPFTable xwpfTable = xwpfDocument.getTableArray(currentTableIndex);
        // 获取表格第一行
        List<XWPFTableCell> tableCells = xwpfTable.getRow(0).getTableCells();
        // 检查是否符合规范
        if (tableCells.size() != 2 || !tableCells.get(0).getText().contains("##{foreachParagraphsStart}##") || tableCells.get(0).getText().trim().length() == 0) {
            throw new InvalidParameterException("文档中第" + (currentTableIndex + 1) + "个表格模板错误，模板表格第一行需要设置2个单元格，" + "第一个单元格存储表格类型(##{foreachParagraphsStart}##)，第二个单元格定义数据源。");
        }

        // 获取数据源标记
        String dataSource = tableCells.get(1).getText();
        // 获取数据源
        List<Map<String, Object>> dataSourceList = (List<Map<String, Object>>) getDataSourceByKey(dataSource, dataMap);

        // 查找段落循环的开始和结束位置
        int beginElementIndex = ++currentElementIndex, beginParagraphIndex = currentParagraphIndex, beginTableIndex = ++currentTableIndex;
        int endElementIndex = 0, endParagraphIndex = 0, endTableIndex = 0;

        // 第一次循环数据并记录段落、表格开始位置和结束位置
        for (int j = 0; j < dataSourceList.size(); j++) {
            endElementIndex = beginElementIndex;
            endParagraphIndex = beginParagraphIndex;
            endTableIndex = beginTableIndex;
            boolean isEnd = false;
            while (!isEnd) {
                IBodyElement element = bodyElement.getBody().getBodyElements().get(endElementIndex);
                switch (element.getElementType()) {
                    case PARAGRAPH:
                        endParagraphIndex = disposeParagraph(bodyElement, endParagraphIndex, dataSourceList.get(j));
                        endElementIndex++;
                        break;
                    case TABLE:
                        XWPFTable xwpfTable1 = xwpfDocument.getTableArray(endTableIndex);
                        if (xwpfTable1.getText().contains("##{foreachParagraphsEnd}##")) {
                            isEnd = true;
                            endTableIndex++;
                            endElementIndex++;
                        } else if (xwpfTable1.getText().contains("##{foreachParagraphsStart}##")) {
                            int[] arrayIndex = disposeLoopsParagraphs(bodyElement, endElementIndex, endParagraphIndex, endTableIndex, dataSourceList.get(j));
                            endElementIndex = arrayIndex[0];
                            endParagraphIndex = arrayIndex[1];
                            endTableIndex = arrayIndex[2];
                        } else {
                            endTableIndex = disposeTable(bodyElement, endTableIndex, dataSourceList.get(j), dataSourceList.get(j));
                            endElementIndex++;
                        }
                        break;
                    default:
                }

            }
        }
        // 返回结束标记位
        return new int[] {endElementIndex, endParagraphIndex, endTableIndex};
    }

    /**
     * 处理文档
     * fixme：未实现
     */
    private void disposeFooters() {

    }

    /**
     * @函数名称: replaceDocument
     * @功能描述: 根据 dataMap 对 word 文件中的标签进行替换;
     * 对于需要替换的普通标签数据标签（不需要循环）必须在 dataMap 中存储一个 key 为 parametersMap 的 map，
     * 来存储这些不需要循环生成的数据，比如：表头信息，日期，制表人等。
     * 对于需要循环生成的表格数据 key 自定义，value 为 ArrayList<Map<String, String>>
     * @输入参数: dataMap 替换模板的数据集
     * @返回值: void
     * @作者: 王艺 <wangyi@fengtaisec.com>
     * @日期: 2022-05-11 11:14:26
     * @修改记录:
     */
    public void replaceDocument() throws Exception {

        // 处理文档页眉
        disposeHeaders();

        // 处理文档页脚
        disposeFooters();

        // 处理段落和表格
        disposeBodyElements();

    }

    /**
     * 处理表格
     *
     * @param bodyElement 文档对象
     * @param dataMap     表格数据对象
     * @param baseData    全局数据对象
     * @return 表格索引
     * @throws Exception
     */
    private int disposeTable(IBodyElement bodyElement, int currentIndex, Map<String, Object> dataMap, Map<String, Object> baseData) throws Exception {
        // 获取表格列表
        XWPFTable xwpfTable = bodyElement.getBody().getTables().get(currentIndex);

        if (xwpfTable != null) {
            // 获取到模板表格第一行，用来判断表格类型
            List<XWPFTableCell> xwpfTableCellList = xwpfTable.getRows().get(0).getTableCells();
            // 表格中的所有文本
            String xwpfTableText = xwpfTable.getText();
            // 段落循环
            if (xwpfTableText.contains("##{foreach")) {
                currentElementIndex++;

                // 查找到 ##{foreach 标签，该表格需要处理循环
                if (xwpfTableCellList.size() != 2 || !xwpfTableCellList.get(0).getText().contains("##{foreach") || xwpfTableCellList.get(0).getText().trim().length() == 0) {
                    throw new InvalidParameterException("文档中第" + (currentIndex + 1) + "个表格模板错误，模板表格第一行需要设置2个单元格，" + "第一个单元格存储表格类型(##{foreachTable}## 或者 " + "##{foreachTableRow}##)，第二个单元格定义数据源。");
                }

                /**
                 * 获取到循环标记
                 * tableType[##{foreachTable}##][##{foreachTableRow}##]
                 * dataSource
                 */
                String tableType = xwpfTableCellList.get(0).getText();
                String dataSource = xwpfTableCellList.get(1).getText();

                log.info("读取到数据源：" + dataSource);
                @SuppressWarnings("unchecked")
                List<Map<String, Object>> tableDataList = (List<Map<String, Object>>) getDataSourceByKey(dataSource, dataMap);

                switch (tableType) {
                    case "##{foreachTable}##":
                        log.info("循环生成表格");
                        addTableInDocFooter(xwpfTable, tableDataList, baseData, ReplaceTypes.TABLE_EXTERNAL_LOOPS);
                        break;
                    case "##{foreachTableRow}##":
                        log.info("循环生成表格内部的行");
                        addTableInDocFooter(xwpfTable, tableDataList, baseData, ReplaceTypes.TABLE_INTERNAL_LOOPS);
                        break;
                }
            } else if (xwpfTableText.contains("##{addChart}##")) {
                currentElementIndex++;
                for (XWPFTableRow row : xwpfTable.getRows()) {
                    if (!row.getCell(0).getText().contains("##{addChart}##"))
                        continue;
                    if (row.getTableCells().size() != 2 || !row.getCell(0).getText().contains("##{addChart") || row.getCell(0).getText().trim().length() == 0) {
                        throw new InvalidParameterException("文档中第" + (currentIndex + 1) + "个表格模板错误，模板表格第一行需要设置3个单元格，" + "第一个单元格存储表格类型(##{addChart}##)，第二个单元格定义数据源。");
                    }
                    String dataSource = row.getCell(1).getText();
                    log.info("读取到数据源：" + dataSource);
                    ChartData chartData = null;
                    Object dataSourceByKey = getDataSourceByKey(dataSource, dataMap);
                    if (dataSourceByKey instanceof ChartData) {
                        chartData = (ChartData) dataSourceByKey;
                    } else {
                        chartData = JSONObject.parseObject(dataSourceByKey.toString(), ChartData.class);
                    }
                    addChartInDocFooter(chartData);
                }
            } else if (xwpfTableText.contains("{")) {
                currentElementIndex++;
                // 没有查找到 ##{foreach 标签，查找到了普通替换数据的 {} 标签，该表格只需要简单替换
                addTableInDocFooter(xwpfTable, null, baseData, ReplaceTypes.IS_REPLACE);
            } else {
                currentElementIndex++;
                // 没有查找到任何标签，该表格是一个静态表格，仅需要复制一个即可。
                addTableInDocFooter(xwpfTable, null, null, ReplaceTypes.IS_COPY);
            }
            return ++currentIndex;
        }
        return ++currentIndex;
    }

    /**
     * 根据 模板表格 和 数据 list 在 word 文档末尾生成表格
     *
     * @param templateTable 模板段落
     * @param list          循环数据集
     * @param parametersMap 模板数据源
     * @param replaceTypes  替换类型
     */
    public void addTableInDocFooter(XWPFTable templateTable
            , List<Map<String, Object>> list
            , Map<String, Object> parametersMap
            , ReplaceTypes replaceTypes) {
        List<XWPFTableRow> templateTableRows;
        XWPFTable newCreateTable;
        switch (replaceTypes) {
            case IS_COPY:
                // 获取模板表格所有行
                templateTableRows = templateTable.getRows();
                // 创建新表格，默认一行一列
                newCreateTable = copyTableStyle(templateTable);

                for (int i = 0; i < templateTableRows.size(); i++) {
                    XWPFTableRow newCreateRow = newCreateTable.createRow();
                    // 复制模板行文本和样式到新行
                    copyTableRowStyle(newCreateRow, templateTableRows.get(i));
                }
                // 移除多出来的第一行
                newCreateTable.removeRow(0);

                // 添加回车换行
                xwpfDocument.createParagraph();
                break;
            case IS_REPLACE:
                templateTableRows = templateTable.getRows();
                newCreateTable = copyTableStyle(templateTable);
                for (int i = 0; i < templateTableRows.size(); i++) {
                    XWPFTableRow newCreateRow = newCreateTable.createRow();
                    // 复制模板行文本和样式到新行
                    copyTableRowStyle(newCreateRow, templateTableRows.get(i));
                }
                // 移除多出来的第一行
                newCreateTable.removeRow(0);
                // 添加回车换行
                xwpfDocument.createParagraph();
                replaceTable(newCreateTable, parametersMap);
                break;
            // 表格整体循环
            case TABLE_EXTERNAL_LOOPS:
                for (Map<String, Object> map : list) {
                    // 获取模板表格所有行
                    templateTableRows = templateTable.getRows();
                    // 创建新表格,默认一行一列
                    newCreateTable = copyTableStyle(templateTable);
                    for (int i = 1; i < templateTableRows.size(); i++) {
                        XWPFTableRow newCreateRow = newCreateTable.createRow();
                        // 复制模板行文本和样式到新行
                        copyTableRowStyle(newCreateRow, templateTableRows.get(i));
                    }
                    // 移除多出来的第一行
                    newCreateTable.removeRow(0);
                    // 添加回车换行
                    xwpfDocument.createParagraph();
                    // 替换标签
                    replaceTable(newCreateTable, map);
                }
                break;
            case TABLE_INTERNAL_LOOPS:
                // 获取模板表格所有行
                List<XWPFTableRow> TempTableRows = templateTable.getRows();
                // 创建新表格,默认一行一列
                newCreateTable = copyTableStyle(templateTable);
                // 标签行 indexs
                int tagRowsIndex = 0;
                for (int i = 0, size = TempTableRows.size(); i < size; i++) {
                    // 获取到表格行的第一个单元格
                    String rowText = TempTableRows.get(i).getCell(0).getText();
                    if (rowText.indexOf("##{foreachRows}##") > -1) {
                        tagRowsIndex = i;
                        break;
                    }
                }

                /** 复制模板行和标签行之前的行 */
                for (int i = 1; i < tagRowsIndex; i++) {
                    XWPFTableRow newCreateRow = newCreateTable.createRow();
                    // 复制行
                    copyTableRowStyle(newCreateRow, TempTableRows.get(i));
                    // 处理不循环标签的替换
                    replaceTableRow(newCreateRow, parametersMap);
                }

                /** 循环生成模板行 */
                // 获取到模板行
                XWPFTableRow tempRow = TempTableRows.get(tagRowsIndex + 1);
                for (int i = 0; i < list.size(); i++) {
                    XWPFTableRow newCreateRow = newCreateTable.createRow();
                    // 复制模板行
                    copyTableRowStyle(newCreateRow, tempRow);
                    // 处理标签替换
                    replaceTableRow(newCreateRow, list.get(i));
                }

                /** 复制模板行和标签行之后的行 */
                for (int i = tagRowsIndex + 2; i < TempTableRows.size(); i++) {
                    XWPFTableRow newCreateRow = newCreateTable.createRow();
                    // 复制行
                    copyTableRowStyle(newCreateRow, TempTableRows.get(i));
                    // 处理不循环标签的替换
                    replaceTableRow(newCreateRow, parametersMap);
                }
                // 移除多出来的第一行
                newCreateTable.removeRow(0);
                // 添加回车换行
                xwpfDocument.createParagraph();
                break;
            default:
                log.info("错误的表格处理方式");
        }
    }

    /**
     * 创建新的表格并同步表格样式
     *
     * @param srcXwpfTable 模板表格
     * @return
     */
    private XWPFTable copyTableStyle(XWPFTable srcXwpfTable) {
        XWPFTable desXwpfTable = xwpfDocument.createTable();
        desXwpfTable.getCTTbl().setTblPr(srcXwpfTable.getCTTbl().getTblPr());
        return desXwpfTable;
    }

    /**
     * 复制表格行 XWPFTableRow 格式
     *
     * @param desXwpfTableRow 待修改格式的 XWPFTableRow
     * @param srcXwpfTableRow 模板 XWPFTableRow
     */
    private void copyTableRowStyle(XWPFTableRow desXwpfTableRow, XWPFTableRow srcXwpfTableRow) {
        // 模板行的列数
        int srcXwpfTableRowCellSize = srcXwpfTableRow.getTableCells().size();
        for (int i = 0; i < srcXwpfTableRowCellSize - 1; i++) {
            // 为新添加的行添加与模板表格对应行行相同个数的单元格
            desXwpfTableRow.addNewTableCell();
        }
        // 复制样式
        desXwpfTableRow.getCtRow().setTrPr(srcXwpfTableRow.getCtRow().getTrPr());
        // 复制单元格
        for (int i = 0; i < desXwpfTableRow.getTableCells().size(); i++) {
            copyTableCellStyle(desXwpfTableRow.getCell(i), srcXwpfTableRow.getCell(i));
        }
    }

    /**
     * 复制单元格 XWPFTableCell 格式
     *
     * @param desXwpfTableCell 目标模板单元格
     * @param srcXwpfParagraph 源模板单元格
     */
    private void copyTableCellStyle(XWPFTableCell desXwpfTableCell, XWPFTableCell srcXwpfParagraph) {
        // 列属性
        desXwpfTableCell.getCTTc().setTcPr(srcXwpfParagraph.getCTTc().getTcPr());
        // 删除目标 targetCell 所有文本段落
        for (int i = 0; i < desXwpfTableCell.getParagraphs().size(); i++) {
            desXwpfTableCell.removeParagraph(i);
        }
        // 添加新文本段落
        for (XWPFParagraph srcParagraph : srcXwpfParagraph.getParagraphs()) {
            XWPFParagraph desParagraph = desXwpfTableCell.addParagraph();
            copyParagraphStyle(desParagraph, srcParagraph);
        }
    }

    /**
     * 根据参数 parametersMap 对表格的一行进行标签的替换
     *
     * @param xwpfTableRow 表格行
     * @param dataSource   数据源
     */
    public void replaceTableRow(XWPFTableRow xwpfTableRow, Map<String, Object> dataSource) {

        List<XWPFTableCell> xwpfTableCellList = xwpfTableRow.getTableCells();
        for (XWPFTableCell xwpfTableCell : xwpfTableCellList) {
            List<XWPFParagraph> xwpfParagraphList = xwpfTableCell.getParagraphs();
            for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                replaceParagraph(xwpfParagraph, dataSource);
            }
        }

    }

    /**
     * 根据数据源替换表格中的 {key} 标签
     *
     * @param xwpfTable  表格
     * @param dataSource 数据源
     */
    public void replaceTable(XWPFTable xwpfTable, Map<String, Object> dataSource) {
        List<XWPFTableRow> xwpfTableRowList = xwpfTable.getRows();

        // 遍历表格每一行
        for (XWPFTableRow xwpfTableRow : xwpfTableRowList) {
            List<XWPFTableCell> xwpfTableCellList = xwpfTableRow.getTableCells();
            // 遍历行内每一个单元格
            for (XWPFTableCell xwpfTableCell : xwpfTableCellList) {
                List<XWPFParagraph> xwpfParagraphList = xwpfTableCell.getParagraphs();
                for (XWPFParagraph xwpfParagraph : xwpfParagraphList) {
                    // 替换段落内容
                    replaceParagraph(xwpfParagraph, dataSource);
                }
            }
        }
    }

    /**
     * 根据模板内容在文档末尾生成图表
     *
     * @param sourceData 模板数据源
     * @throws Exception
     */
    public void addChartInDocFooter(ChartData sourceData) throws Exception {
        if (sourceData.categories.categories.length == 0 || sourceData.seriesData.series.length == 0) {
            XWPFParagraph paragraph = xwpfDocument.createParagraph();
            XWPFRun xwpfRun = paragraph.insertNewRun(0);
            xwpfRun.setText(sourceData.title + "暂无数据");
            return;
        }
        // create a histogram
        XWPFChart chart = xwpfDocument.createChart(14 * Units.EMU_PER_CENTIMETER, 10 * Units.EMU_PER_CENTIMETER);
        // set title
        chart.setTitleText(sourceData.title);
        chart.setTitleOverlay(false);
        // create legend
        XDDFChartLegend legend = chart.getOrAddLegend();
        legend.setPosition(LegendPosition.BOTTOM);
        legend.setOverlay(false);

        XDDFCategoryAxis categoryAxis = null;
        XDDFValueAxis valueAxis = null;
        if (ChartTypes.PIE != sourceData.chartType && ChartTypes.PIE3D != sourceData.chartType) {
            // create axis
            categoryAxis = chart.createCategoryAxis(AxisPosition.BOTTOM);
            categoryAxis.setTitle(sourceData.categories.name);
            categoryAxis.setCrosses(AxisCrosses.AUTO_ZERO);

            valueAxis = chart.createValueAxis(AxisPosition.LEFT);
            valueAxis.setTitle(sourceData.seriesData.name);
            valueAxis.setCrosses(AxisCrosses.AUTO_ZERO);
            // Set AxisCrossBetween, so the left axis crosses the category axis between the categories.
            // Else first and last category is exactly on cross points and the bars are only half visible.
            valueAxis.setCrossBetween(AxisCrossBetween.BETWEEN);
        }

        // create chart data
        XDDFChartData data = chart.createData(sourceData.chartType, categoryAxis, valueAxis);

        switch (sourceData.chartType) {
            case BAR:
                processBarData(chart, (XDDFBarChartData) data, sourceData);
                break;
            case LINE:
                processLineData(chart, (XDDFLineChartData) data, sourceData);
                break;
            case PIE:
                processPieData(chart, (XDDFPieChartData) data, sourceData);
                break;
            default:
        }
        // plot chart data
        chart.plot(data);
    }

    /**
     * 处理饼状图数据
     *
     * @param chartData  图表数据
     * @param sourceData 模板数据源
     */
    private void processPieData(XWPFChart chart, XDDFPieChartData chartData, ChartData sourceData) {
        chartData.setVaryColors(true);

        // create data sources
        XDDFDataSource<String> categoryDataSource = XDDFDataSourcesFactory.fromArray(sourceData.categories.categories);

        for (int i = 0; i < sourceData.seriesData.series.length; i++) {
            XDDFNumericalDataSource valueData = XDDFDataSourcesFactory.fromArray(sourceData.seriesData.series[i].values);
            XDDFPieChartData.Series series = (XDDFPieChartData.Series) chartData.addSeries(categoryDataSource, valueData);
            series.setTitle(sourceData.seriesData.series[i].name, null);
            series.setShowLeaderLines(true);
        }
    }

    /**
     * 处理折线图数据
     *
     * @param chart      图表对象
     * @param chartData  图表数据
     * @param sourceData 模板数据源
     * @throws Exception
     */
    private void processLineData(XWPFChart chart, XDDFLineChartData chartData, ChartData sourceData) throws Exception {
        chartData.setVaryColors(true);
        chartData.setGrouping(Grouping.STANDARD);

        // create data sources
        String categoryDataRange = chart.formatRange(new CellRangeAddress(1, sourceData.categories.categories.length, 0, 0));
        XDDFDataSource<String> categoryDataSource = XDDFDataSourcesFactory.fromArray(sourceData.categories.categories, categoryDataRange, 0);

        for (int i = 0; i < sourceData.seriesData.series.length; i++) {
            String seriesDataRange = chart.formatRange(new CellRangeAddress(1, sourceData.seriesData.series[i].values.length, i + 1, i + 1));
            XDDFNumericalDataSource<Double> valueData = XDDFDataSourcesFactory.fromArray(sourceData.seriesData.series[i].values, seriesDataRange, i + 1);
            XDDFLineChartData.Series series = (XDDFLineChartData.Series) chartData.addSeries(categoryDataSource, valueData);
            series.setTitle(sourceData.seriesData.series[i].name, setTitleInDataSheet(chart, sourceData.seriesData.series[i].name, i + 1));
        }
    }

    /**
     * 处理条形图数据
     *
     * @param chart      图表对象
     * @param chartData  图表数据
     * @param sourceData 模板数据源
     * @throws Exception
     */
    private void processBarData(XWPFChart chart, XDDFBarChartData chartData, ChartData sourceData) throws Exception {
        chartData.setBarDirection(BarDirection.BAR);
        chartData.setBarGrouping(BarGrouping.STANDARD);
        chartData.setVaryColors(true);

        // create data sources
        XDDFDataSource<String> categoryDataSource = XDDFDataSourcesFactory.fromArray(sourceData.categories.categories);

        for (int i = 0; i < sourceData.seriesData.series.length; i++) {
            XDDFNumericalDataSource<Double> valueData = XDDFDataSourcesFactory.fromArray(sourceData.seriesData.series[i].values);
            XDDFBarChartData.Series series = (XDDFBarChartData.Series) chartData.addSeries(categoryDataSource, valueData);
            series.setTitle(sourceData.seriesData.series[i].name, null);
            series.setShowLeaderLines(true);
        }
    }

    /**
     * Methode to set title in the data sheet without creating a Table but using the sheet data only.
     * Creating a Table is not really necessary.
     *
     * @param chart  图表对象
     * @param title  数据系列标题
     * @param column 列
     * @return
     * @throws Exception
     */
    private CellReference setTitleInDataSheet(XWPFChart chart, String title, int column) throws Exception {
        XSSFWorkbook workbook = chart.getWorkbook();
        XSSFSheet sheet = workbook.getSheetAt(0);
        XSSFRow row = sheet.getRow(0);
        if (row == null) row = sheet.createRow(0);
        XSSFCell cell = row.getCell(column);
        if (cell == null) cell = row.createCell(column);
        cell.setCellValue(title);
        return new CellReference(sheet.getSheetName(), 0, column, true, true);
    }

    /**
     * 处理段落
     *
     * @return
     */
    private int disposeParagraph(IBodyElement bodyElement, int currentIndex, Map<String, Object> baseData) {
        log.info("获取到段落");
        XWPFParagraph xwpfParagraph = bodyElement.getBody().getParagraphArray(currentIndex);

        if (!xwpfParagraph.isEmpty()) {
            addParagraphInDocFooter(xwpfParagraph, null, baseData, ReplaceTypes.IS_NOT_LOOPS);
            return ++currentIndex;
        } else {
            return ++currentIndex;
        }
    }

    /**
     * 根据模板段落和数据在文档末尾生成段落
     *
     * @param srcXwpfParagraph 模板段落
     * @param list             循环数据集
     * @param baseData         模板数据源
     * @param replaceTypes     模板处理逻辑
     */
    public void addParagraphInDocFooter(XWPFParagraph srcXwpfParagraph, List<Map<String, Object>> list, Map<String, Object> baseData, ReplaceTypes replaceTypes) {
        switch (replaceTypes) {
            case IS_LOOPS:
                // 暂无实现
                break;
            case IS_NOT_LOOPS:
                XWPFParagraph desXwpfParagraph = xwpfDocument.createParagraph();
                copyParagraphStyle(desXwpfParagraph, srcXwpfParagraph);
                replaceParagraph(desXwpfParagraph, baseData);
                break;
            default:
        }
    }

    /**
     * 复制文本段落 XWPFParagraph 格式
     *
     * @param desXwpfParagraph 目标模板段落段落
     * @param srcXwpfParagraph 源模板段落节点
     */
    private void copyParagraphStyle(XWPFParagraph desXwpfParagraph, XWPFParagraph srcXwpfParagraph) {
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
     * 根据数据源替换段落元素内的 {**} 标签
     *
     * @param xwpfParagraph 段落对象
     * @param dataSource    模板数据源
     */
    public void replaceParagraph(XWPFParagraph xwpfParagraph, Map<String, Object> dataSource) {
        List<XWPFRun> xwpfRuns = xwpfParagraph.getRuns();
        String xwpfParagraphText = xwpfParagraph.getText();

        //正则匹配字符串{****}
        String regEx = "\\{.+?\\}";
        Pattern pattern = Pattern.compile(regEx);
        Matcher matcher = pattern.matcher(xwpfParagraphText);
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
                    newRun.setText(getDataSourceByKey(key.toString(), dataSource).toString());
                } else {
                    // 该run标签为**{**}** 或者 **{**} 或者{**}**，替换key后，还需要加上原始key前后的文本
                    // 设置文本
                    key.append(beginXwpfRunText.substring(beginXwpfRunText.indexOf("{") + 1, beginXwpfRunText.indexOf("}")));
                    String textString = beginXwpfRunText.substring(0, beginTagIndex) + getDataSourceByKey(key.toString(), dataSource) + beginXwpfRunText.substring(endTagIndex + 1);
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
                    insertNewRun.setText(getDataSourceByKey(key.toString(), dataSource).toString());
                    // 移除原始的run
                    xwpfParagraph.removeRun(beginRunIndex + 1);
                } else {
                    // 该run标签为**{**或者 {** ，替换key后，还需要加上原始key前的文本
                    XWPFRun insertNewRun = xwpfParagraph.insertNewRun(beginRunIndex);
                    insertNewRun.getCTR().setRPr(beginRun.getCTR().getRPr());
                    // 设置文本
                    String textString = beginRunText.substring(0, beginRunText.indexOf("{")) + getDataSourceByKey(key.toString(), dataSource);
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
            replaceParagraph(xwpfParagraph, dataSource);
        }
    }

    /**
     * 根据 key 获取数据源信息
     *
     * @param key        数据键
     * @param dataSource 数据源
     * @return
     */
    private Object getDataSourceByKey(String key, Map<String, Object> dataSource) {
        Object result = "";

        if (key.isEmpty()) {
            return result;
        }

        if (key.contains(".")) {
            dataSource = globalData;
        }
        result = parseKey(key.split("\\."), dataSource);
        if (result == null) {
            log.error("数据源缺失：" + key);
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
}

