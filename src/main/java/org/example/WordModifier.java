////////////////////////package org.example;
////////////////////////
////////////////////////import com.aspose.words.SaveFormat;
////////////////////////import org.apache.poi.xwpf.usermodel.*;
////////////////////////
////////////////////////import java.io.File;
////////////////////////import java.io.FileInputStream;
////////////////////////import java.io.FileOutputStream;
////////////////////////import java.io.IOException;
////////////////////////import java.util.HashMap;
////////////////////////import java.util.List;
////////////////////////import java.util.Map;
////////////////////////
////////////////////////
////////////////////////
////////////////////////public class WordModifier {
////////////////////////    public static void main(String[] args) {
////////////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx"; // 配置文档的路径
////////////////////////
////////////////////////        File folder = new File("Z:\\Desktop\\测试\\in"); // 文件夹的路径
////////////////////////        File[] listOfFiles = folder.listFiles(); // 获取文件夹中的所有文件
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////        for (File file : listOfFiles) {
////////////////////////            if (file.isFile()) { // 检查文件是否是Word文档
////////////////////////                String sourceFile = file.getAbsolutePath(); // 获取文件的绝对路径
////////////////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName(); // 输出文档的路径
////////////////////////
////////////////////////                // 如果文件是.doc文件，将其转换为.docx
////////////////////////                if (file.getName().endsWith(".doc")) {
////////////////////////
////////////////////////                    String docxFile = sourceFile.replace(".doc", ".docx");
////////////////////////                                        // 使用Aspose Words的Document类进行转换
////////////////////////                                        com.aspose.words.Document doc = null;
////////////////////////                                        try {
////////////////////////                                                doc = new com.aspose.words.Document(sourceFile);
////////////////////////                                        } catch (Exception e) {
////////////////////////                                                throw new RuntimeException(e);
////////////////////////                                        }
////////////////////////                                        try {
////////////////////////                                                doc.save(docxFile, SaveFormat.DOCX);
////////////////////////                                        } catch (Exception e) {
////////////////////////                                                throw new RuntimeException(e);
////////////////////////                                        }
////////////////////////                                        sourceFile = docxFile; // 更新源文件路径为.docx文件
//////////////////////////                    // 加载源 PDF 文件
//////////////////////////                    Converter converter = new Converter(sourceFile);
//////////////////////////
//////////////////////////// 设置转换选项
//////////////////////////                    WordProcessingConvertOptions convertOptions =
//////////////////////////                            new WordProcessingConvertOptions();
//////////////////////////
//////////////////////////// 将 PDF 转换为 DOCX
//////////////////////////                    converter.convert(docxFile, convertOptions);
////////////////////////                }
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////                // 确保源文件是.docx文件
////////////////////////                if (!sourceFile.endsWith(".docx")) {
////////////////////////                    continue;
////////////////////////                }
////////////////////////                Map<String, String> configMap = new HashMap<>(); // 创建一个映射来存储配置规则
////////////////////////
////////////////////////                // 读取配置文档
////////////////////////                try (FileInputStream configFis = new FileInputStream(configFile);
////////////////////////                     XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////////////////
////////////////////////                    List<XWPFTable> configTables = configDoc.getTables(); // 获取配置文档中的所有表格
////////////////////////
////////////////////////                    // 遍历配置文档中的每个表格
////////////////////////                    for (XWPFTable table : configTables) {
////////////////////////                        // 遍历每个表格中的行
////////////////////////                        for (XWPFTableRow row : table.getRows()) {
////////////////////////                            // 假设每行的第一个单元格包含标签，第二个单元格包含更新的值
////////////////////////                            if (row.getTableCells().size() >= 2) {
////////////////////////                                XWPFTableCell labelCell = row.getTableCells().get(0);
////////////////////////                                XWPFTableCell valueCell = row.getTableCells().get(1);
////////////////////////                                String key = labelCell.getText().trim();
////////////////////////                                String value = valueCell.getText().trim();
////////////////////////                                configMap.put(key, value); // 将键值对添加到映射中
////////////////////////                            }
////////////////////////                        }
////////////////////////                    }
////////////////////////                } catch (IOException e) {
////////////////////////                    e.printStackTrace();
////////////////////////                }
////////////////////////
////////////////////////                // 打印配置映射，调试用
////////////////////////                System.out.println("配置映射：");
////////////////////////                for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////////////////                                        System.out.println(entry.getKey() + " => " + entry.getValue());
////////////////////////                }
////////////////////////
////////////////////////                // 读取和修改目标文档
////////////////////////                try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////////////////                     FileOutputStream fos = new FileOutputStream(outputFile);
////////////////////////                     XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////////////////
////////////////////////                    List<XWPFTable> tables = document.getTables(); // 获取源文档中的所有表格
////////////////////////
////////////////////////                    // 遍历源文档中的每个表格
////////////////////////                    for (XWPFTable table : tables) {
////////////////////////                        // 遍历每个表格中的行
////////////////////////                        for (XWPFTableRow row : table.getRows()) {
////////////////////////                            // 遍历行中的每个单元格
////////////////////////                            for (int i = 0; i < row.getTableCells().size(); i++) {
////////////////////////                                XWPFTableCell cell = row.getTableCells().get(i);
////////////////////////                                String text = cell.getText().replaceAll("\\s+", ""); // 获取单元格中的文本
////////////////////////
////////////////////////                                // 打印调试信息
//////////////////////////                                                                System.out.println("表格单元格文本：" + text);
////////////////////////
////////////////////////                                // 如果单元格的文本存在于配置规则的映射中
////////////////////////                                if (configMap.containsKey(text)) {
////////////////////////                                    String newValue = configMap.get(text);
////////////////////////                                    // 检查并更新下一个单元格（如果存在和有新值）
////////////////////////                                    if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////////////////                                        XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////////////////
////////////////////////                                        // 清空单元格
////////////////////////                                        nextCell.removeParagraph(0);
////////////////////////
////////////////////////                                        // 添加新的内容
////////////////////////                                        XWPFParagraph p = nextCell.addParagraph();
////////////////////////                                        p.setAlignment(ParagraphAlignment.CENTER); // 设置段落为居中对齐
////////////////////////                                        XWPFRun r = p.createRun();
////////////////////////                                        r.setText(newValue);
////////////////////////
////////////////////////                                        nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); // 设置单元格内容为水平居中
////////////////////////                                    }
////////////////////////                                }
////////////////////////                            }
////////////////////////                        }
////////////////////////                    }
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////                    // 遍历文档中的所有段落
////////////////////////                    for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////////////////                        List<XWPFRun> runs = paragraph.getRuns();
////////////////////////                        for (int i = 0; i < runs.size(); i++) {
////////////////////////                            XWPFRun run = runs.get(i);
////////////////////////                            String text = run.getText(0);
////////////////////////                            if (text != null) {
////////////////////////                                for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////////////                                    String key = entry.getKey();
////////////////////////                                    String value = entry.getValue();
////////////////////////                                    if(key.equals(text.trim())){
////////////////////////                                        System.out.println(key+" **   "+text.trim()+"      是否相等  "+key.equals(text.trim()));
////////////////////////                                    }
////////////////////////                                    if (text.trim().equals(key)) { // 检查文本是否包含键和一个空格和一个冒号
////////////////////////                                        // 修改后面的文本
////////////////////////                                        int j = i + 1;
////////////////////////                                        while (j < runs.size()) {
////////////////////////                                            XWPFRun nextRun = runs.get(j);
////////////////////////                                            String nextText = nextRun.getText(0);
////////////////////////                                            System.out.println("########");
////////////////////////                                            System.out.println(nextText);
////////////////////////                                            if (nextText != null && !nextText.contains(":")) {
////////////////////////                                                // 创建新的 run 并设置文本
////////////////////////                                                if (j+1 <= runs.size()) {
////////////////////////                                                    XWPFRun newRun = paragraph.insertNewRun(j+1);
////////////////////////                                                    newRun.setText(value+"   ");
////////////////////////
////////////////////////                                                    // 给新的 run 添加下划线
////////////////////////                                                    newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////////////////
////////////////////////                                                    // 设置字体和字号
////////////////////////                                                    newRun.setFontFamily("仿宋_GB2312");
////////////////////////                                                    newRun.setFontSize(14); // 四号字体对应的字号大约为14pt
////////////////////////                                                    // 删除旧的 run
////////////////////////                                                    paragraph.removeRun(j);
////////////////////////                                                }
////////////////////////
////////////////////////                                                break;
////////////////////////                                            }
////////////////////////
////////////////////////
////////////////////////
////////////////////////                                            j++;
////////////////////////                                        }
////////////////////////                                        i = j; // 跳过已经修改的运行
////////////////////////                                        break; // 找到一个匹配的键后，就退出循环
////////////////////////                                    }
////////////////////////                                }
////////////////////////                            }
////////////////////////                        }
////////////////////////                    }
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////
////////////////////////                    document.write(fos); // 将修改后的文档写入到输出文件中
////////////////////////                } catch (IOException e) {
////////////////////////                    e.printStackTrace();
////////////////////////                }
////////////////////////
////////////////////////
////////////////////////
////////////////////////            }
////////////////////////        }
////////////////////////    }
////////////////////////
////////////////////////
////////////////////////
////////////////////////}
////////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////package org.example;
//////////////////////
//////////////////////import com.aspose.words.SaveFormat;
//////////////////////import org.apache.poi.xwpf.usermodel.*;
//////////////////////import javax.swing.*;
//////////////////////import javax.swing.table.DefaultTableModel;
//////////////////////import org.jfree.chart.ChartFactory;
//////////////////////import org.jfree.chart.ChartPanel;
//////////////////////import org.jfree.chart.JFreeChart;
//////////////////////import org.jfree.chart.plot.PlotOrientation;
//////////////////////import org.jfree.data.category.DefaultCategoryDataset;
//////////////////////
//////////////////////import java.awt.*;
//////////////////////import java.io.File;
//////////////////////import java.io.FileInputStream;
//////////////////////import java.io.FileOutputStream;
//////////////////////import java.io.IOException;
//////////////////////import java.util.HashMap;
//////////////////////import java.util.List;
//////////////////////import java.util.Map;
//////////////////////
//////////////////////public class WordModifier {
//////////////////////    private static Map<String, String> configMap = new HashMap<>();
//////////////////////    private static DefaultTableModel tableModel = new DefaultTableModel(new Object[]{"文件", "错误类型", "描述"}, 0);
//////////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////////////////    private static DefaultCategoryDataset dataset = new DefaultCategoryDataset();
//////////////////////
//////////////////////    public static void main(String[] args) {
//////////////////////        // 创建并显示主窗口
//////////////////////        JFrame frame = createMainFrame();
//////////////////////        frame.setVisible(true);
//////////////////////
//////////////////////        // 加载配置文件
//////////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////////////////        loadConfigFile(configFile);
//////////////////////
//////////////////////        // 处理文档文件夹中的文件
//////////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////////////////        File[] listOfFiles = folder.listFiles();
//////////////////////
//////////////////////        if (listOfFiles != null) {
//////////////////////            int totalFiles = listOfFiles.length;
//////////////////////            int processedFiles = 0;
//////////////////////            int successCount = 0;
//////////////////////            int failureCount = 0;
//////////////////////
//////////////////////            for (File file : listOfFiles) {
//////////////////////                if (file.isFile()) {
//////////////////////                    String sourceFile = file.getAbsolutePath();
//////////////////////                    String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////////////////
//////////////////////                    // 如果文件是.doc文件，将其转换为.docx
//////////////////////                    if (sourceFile.endsWith(".doc")) {
//////////////////////                        sourceFile = convertDocToDocx(sourceFile);
//////////////////////                    }
//////////////////////
//////////////////////                    // 确保源文件是.docx文件
//////////////////////                    if (!sourceFile.endsWith(".docx")) {
//////////////////////                        continue;
//////////////////////                    }
//////////////////////
//////////////////////                    try {
//////////////////////                        modifyDocument(sourceFile, outputFile);
//////////////////////                        successCount++;
//////////////////////                    } catch (Exception e) {
//////////////////////                        tableModel.addRow(new Object[]{file.getName(), "处理错误", e.getMessage()});
//////////////////////                        failureCount++;
//////////////////////                    }
//////////////////////
//////////////////////                    processedFiles++;
//////////////////////                    updateProgress(processedFiles, totalFiles);
//////////////////////                }
//////////////////////            }
//////////////////////
//////////////////////            dataset.addValue(successCount, "数量", "成功");
//////////////////////            dataset.addValue(failureCount, "数量", "失败");
//////////////////////        }
//////////////////////    }
//////////////////////
//////////////////////    private static void loadConfigFile(String configFile) {
//////////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////////////
//////////////////////            List<XWPFTable> configTables = configDoc.getTables();
//////////////////////
//////////////////////            for (XWPFTable table : configTables) {
//////////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////////                    if (row.getTableCells().size() >= 2) {
//////////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////////////                        String key = labelCell.getText().trim();
//////////////////////                        String value = valueCell.getText().trim();
//////////////////////                        configMap.put(key, value);
//////////////////////                    }
//////////////////////                }
//////////////////////            }
//////////////////////        } catch (IOException e) {
//////////////////////            e.printStackTrace();
//////////////////////        }
//////////////////////    }
//////////////////////
//////////////////////    private static String convertDocToDocx(String sourceFile) {
//////////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////////////////        try {
//////////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////////////////            doc.save(docxFile, SaveFormat.DOCX);
//////////////////////        } catch (Exception e) {
//////////////////////            throw new RuntimeException(e);
//////////////////////        }
//////////////////////        return docxFile;
//////////////////////    }
//////////////////////
//////////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////////////
//////////////////////            List<XWPFTable> tables = document.getTables();
//////////////////////
//////////////////////            for (XWPFTable table : tables) {
//////////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////////////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////////////////
//////////////////////                        if (configMap.containsKey(text)) {
//////////////////////                            String newValue = configMap.get(text);
//////////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////////////                                nextCell.removeParagraph(0);
//////////////////////                                XWPFParagraph p = nextCell.addParagraph();
//////////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////////////////                                XWPFRun r = p.createRun();
//////////////////////                                r.setText(newValue);
//////////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////////////////                            }
//////////////////////                        }
//////////////////////                    }
//////////////////////                }
//////////////////////            }
//////////////////////
//////////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////////////                List<XWPFRun> runs = paragraph.getRuns();
//////////////////////                for (int i = 0; i < runs.size(); i++) {
//////////////////////                    XWPFRun run = runs.get(i);
//////////////////////                    String text = run.getText(0);
//////////////////////                    if (text != null) {
//////////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////////////                            String key = entry.getKey();
//////////////////////                            String value = entry.getValue();
//////////////////////                            if (text.trim().equals(key)) {
//////////////////////                                int j = i + 1;
//////////////////////                                while (j < runs.size()) {
//////////////////////                                    XWPFRun nextRun = runs.get(j);
//////////////////////                                    String nextText = nextRun.getText(0);
//////////////////////                                    if (nextText != null && !nextText.contains(":")) {
//////////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////////////////                                        newRun.setText(value);
//////////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////////////////                                        newRun.setFontSize(14);
//////////////////////                                        paragraph.removeRun(j);
//////////////////////                                        break;
//////////////////////                                    }
//////////////////////                                    j++;
//////////////////////                                }
//////////////////////                                i = j;
//////////////////////                                break;
//////////////////////                            }
//////////////////////                        }
//////////////////////                    }
//////////////////////                }
//////////////////////            }
//////////////////////
//////////////////////            document.write(fos);
//////////////////////        }
//////////////////////    }
//////////////////////
//////////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////////////////        progressBar.setValue(progress);
//////////////////////    }
//////////////////////
//////////////////////    private static JFrame createMainFrame() {
//////////////////////        JFrame frame = new JFrame("文档处理工具");
//////////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////////////////        frame.setSize(800, 600);
//////////////////////
//////////////////////        JPanel panel = new JPanel(new BorderLayout());
//////////////////////        frame.add(panel);
//////////////////////
//////////////////////        JTable table = new JTable(tableModel);
//////////////////////        panel.add(new JScrollPane(table), BorderLayout.NORTH);
//////////////////////
//////////////////////        panel.add(progressBar, BorderLayout.CENTER);
//////////////////////
//////////////////////        JFreeChart barChart = ChartFactory.createBarChart(
//////////////////////                "文档处理统计分析",
//////////////////////                "类别",
//////////////////////                "数量",
//////////////////////                dataset,
//////////////////////                PlotOrientation.VERTICAL,
//////////////////////                true, true, false);
//////////////////////        ChartPanel chartPanel = new ChartPanel(barChart);
//////////////////////        chartPanel.setPreferredSize(new Dimension(800, 400));
//////////////////////        panel.add(chartPanel, BorderLayout.SOUTH);
//////////////////////
//////////////////////        return frame;
//////////////////////    }
//////////////////////}
//////////////////////
//////////////////////
//////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////package org.example;
////////////////////
////////////////////import com.aspose.words.SaveFormat;
////////////////////import org.apache.poi.xwpf.usermodel.*;
////////////////////import javax.swing.*;
////////////////////import javax.swing.table.DefaultTableModel;
////////////////////import org.jfree.chart.ChartFactory;
////////////////////import org.jfree.chart.ChartPanel;
////////////////////import org.jfree.chart.JFreeChart;
////////////////////import org.jfree.chart.plot.PlotOrientation;
////////////////////import org.jfree.data.category.DefaultCategoryDataset;
////////////////////
////////////////////import java.awt.*;
////////////////////import java.awt.event.ActionEvent;
////////////////////import java.awt.event.ActionListener;
////////////////////import java.io.File;
////////////////////import java.io.FileInputStream;
////////////////////import java.io.FileOutputStream;
////////////////////import java.io.IOException;
////////////////////import java.util.HashMap;
////////////////////import java.util.List;
////////////////////import java.util.Map;
////////////////////
////////////////////public class WordModifier {
////////////////////    private static Map<String, String> configMap = new HashMap<>();
////////////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////////////    private static DefaultCategoryDataset dataset = new DefaultCategoryDataset();
////////////////////    private static JFrame frame;
////////////////////    private static JTextField keyField = new JTextField();
////////////////////    private static JTextField valueField = new JTextField();
////////////////////    private static JTable configTable;
////////////////////    private static JTable fileTable;
////////////////////
////////////////////    public static void main(String[] args) {
////////////////////        // 创建并显示主窗口
////////////////////        frame = createMainFrame();
////////////////////        frame.setVisible(true);
////////////////////
////////////////////        // 加载配置文件
////////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////////////        loadConfigFile(configFile);
////////////////////
////////////////////        // 处理文档文件夹中的文件
////////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////////////        File[] listOfFiles = folder.listFiles();
////////////////////
////////////////////        if (listOfFiles != null) {
////////////////////            int totalFiles = listOfFiles.length;
////////////////////            int processedFiles = 0;
////////////////////            int successCount = 0;
////////////////////            int failureCount = 0;
////////////////////
////////////////////            for (File file : listOfFiles) {
////////////////////                if (file.isFile()) {
////////////////////                    String sourceFile = file.getAbsolutePath();
////////////////////                    String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////////////
////////////////////                    // 如果文件是.doc文件，将其转换为.docx
////////////////////                    if (sourceFile.endsWith(".doc")) {
////////////////////                        sourceFile = convertDocToDocx(sourceFile);
////////////////////                    }
////////////////////
////////////////////                    // 确保源文件是.docx文件
////////////////////                    if (!sourceFile.endsWith(".docx")) {
////////////////////                        continue;
////////////////////                    }
////////////////////
////////////////////                    try {
////////////////////                        modifyDocument(sourceFile, outputFile);
////////////////////                        successCount++;
////////////////////                    } catch (Exception e) {
////////////////////                        fileTableModel.addRow(new Object[]{file.getName(), "处理错误", e.getMessage()});
////////////////////                        failureCount++;
////////////////////                    }
////////////////////
////////////////////                    processedFiles++;
////////////////////                    updateProgress(processedFiles, totalFiles);
////////////////////                }
////////////////////            }
////////////////////
////////////////////            dataset.addValue(successCount, "数量", "成功");
////////////////////            dataset.addValue(failureCount, "数量", "失败");
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static void loadConfigFile(String configFile) {
////////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////////////
////////////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////////////
////////////////////            for (XWPFTable table : configTables) {
////////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////////                    if (row.getTableCells().size() >= 2) {
////////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////////////                        String key = labelCell.getText().trim();
////////////////////                        String value = valueCell.getText().trim();
////////////////////                        configMap.put(key, value);
////////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////        } catch (IOException e) {
////////////////////            e.printStackTrace();
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static String convertDocToDocx(String sourceFile) {
////////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////////////        try {
////////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////////////        } catch (Exception e) {
////////////////////            throw new RuntimeException(e);
////////////////////        }
////////////////////        return docxFile;
////////////////////    }
////////////////////
////////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////////////
////////////////////            StringBuilder originalContent = new StringBuilder();
////////////////////            StringBuilder modifiedContent = new StringBuilder();
////////////////////
////////////////////            List<XWPFTable> tables = document.getTables();
////////////////////
////////////////////            for (XWPFTable table : tables) {
////////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////////////                        originalContent.append(text).append(" ");
////////////////////
////////////////////                        if (configMap.containsKey(text)) {
////////////////////                            String newValue = configMap.get(text);
////////////////////                            modifiedContent.append(newValue).append(" ");
////////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////////////                                nextCell.removeParagraph(0);
////////////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////////////                                XWPFRun r = p.createRun();
////////////////////                                r.setText(newValue);
////////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////////////                            }
////////////////////                        }
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////
////////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////////////                for (int i = 0; i < runs.size(); i++) {
////////////////////                    XWPFRun run = runs.get(i);
////////////////////                    String text = run.getText(0);
////////////////////                    if (text != null) {
////////////////////                        originalContent.append(text).append(" ");
////////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////////                            String key = entry.getKey();
////////////////////                            String value = entry.getValue();
////////////////////                            if (text.trim().equals(key)) {
////////////////////                                int j = i + 1;
////////////////////                                while (j < runs.size()) {
////////////////////                                    XWPFRun nextRun = runs.get(j);
////////////////////                                    String nextText = nextRun.getText(0);
////////////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////////////                                        newRun.setText(value);
////////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////////////                                        newRun.setFontSize(14);
////////////////////                                        paragraph.removeRun(j);
////////////////////                                        break;
////////////////////                                    }
////////////////////                                    j++;
////////////////////                                }
////////////////////                                i = j;
////////////////////                                break;
////////////////////                            }
////////////////////                        }
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////
////////////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////////////////            document.write(fos);
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////////////        progressBar.setValue(progress);
////////////////////    }
////////////////////
////////////////////    private static JFrame createMainFrame() {
////////////////////        JFrame frame = new JFrame("文档处理工具");
////////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////////////        frame.setSize(1000, 800);
////////////////////
////////////////////        JPanel panel = new JPanel(new BorderLayout());
////////////////////        frame.add(panel);
////////////////////
////////////////////        JPanel configPanel = new JPanel(new BorderLayout());
////////////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////////////////        configTable = new JTable(configTableModel);
////////////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////////////////
////////////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////////////////        configInputPanel.add(new JLabel("Key:"));
////////////////////        configInputPanel.add(keyField);
////////////////////        configInputPanel.add(new JLabel("Value:"));
////////////////////        configInputPanel.add(valueField);
////////////////////
////////////////////        JButton addButton = new JButton("添加/更新");
////////////////////        addButton.addActionListener(new ActionListener() {
////////////////////            @Override
////////////////////            public void actionPerformed(ActionEvent e) {
////////////////////                String key = keyField.getText().trim();
////////////////////                String value = valueField.getText().trim();
////////////////////                if (!key.isEmpty() && !value.isEmpty()) {
////////////////////                    configMap.put(key, value);
////////////////////                    boolean keyExists = false;
////////////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////////////////                            configTableModel.setValueAt(value, i, 1);
////////////////////                            keyExists = true;
////////////////////                            break;
////////////////////                        }
////////////////////                    }
////////////////////                    if (!keyExists) {
////////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////        });
////////////////////        configInputPanel.add(addButton);
////////////////////
////////////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////////////////
////////////////////        JPanel filePanel = new JPanel(new BorderLayout());
////////////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果"));
////////////////////        fileTable = new JTable(fileTableModel);
////////////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////////////////
////////////////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////////////////
////////////////////        JPanel chartPanel = new JPanel(new BorderLayout());
////////////////////        chartPanel.setBorder(BorderFactory.createTitledBorder("文档处理统计分析"));
////////////////////        JFreeChart barChart = ChartFactory.createBarChart(
////////////////////                "文档处理统计分析",
////////////////////                "类别",
////////////////////                "数量",
////////////////////                dataset,
////////////////////                PlotOrientation.VERTICAL,
////////////////////                true, true, false);
////////////////////        ChartPanel chartPanelInner = new ChartPanel(barChart);
////////////////////        chartPanel.add(chartPanelInner, BorderLayout.CENTER);
////////////////////
////////////////////        panel.add(configPanel, BorderLayout.NORTH);
////////////////////        panel.add(filePanel, BorderLayout.CENTER);
////////////////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////////////////        panel.add(chartPanel, BorderLayout.EAST);
////////////////////
////////////////////        return frame;
////////////////////    }
////////////////////}
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////package org.example;
//////////////////
//////////////////import com.aspose.words.SaveFormat;
//////////////////import org.apache.poi.xwpf.usermodel.*;
//////////////////import javax.swing.*;
//////////////////import javax.swing.table.DefaultTableModel;
//////////////////import java.awt.*;
//////////////////import java.awt.event.ActionEvent;
//////////////////import java.awt.event.ActionListener;
//////////////////import java.io.File;
//////////////////import java.io.FileInputStream;
//////////////////import java.io.FileOutputStream;
//////////////////import java.io.IOException;
//////////////////import java.util.HashMap;
//////////////////import java.util.List;
//////////////////import java.util.Map;
//////////////////
//////////////////public class WordModifier {
//////////////////    private static Map<String, String> configMap = new HashMap<>();
//////////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////////////    private static JFrame frame;
//////////////////    private static JTextField keyField = new JTextField();
//////////////////    private static JTextField valueField = new JTextField();
//////////////////    private static JTable configTable;
//////////////////    private static JTable fileTable;
//////////////////    private static JTextArea statsTextArea = new JTextArea();
//////////////////
//////////////////    public static void main(String[] args) {
//////////////////        // 创建并显示主窗口
//////////////////        frame = createMainFrame();
//////////////////        frame.setVisible(true);
//////////////////
//////////////////        // 加载配置文件
//////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////////////        loadConfigFile(configFile);
//////////////////    }
//////////////////
//////////////////    private static void loadConfigFile(String configFile) {
//////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////////
//////////////////            List<XWPFTable> configTables = configDoc.getTables();
//////////////////
//////////////////            for (XWPFTable table : configTables) {
//////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////                    if (row.getTableCells().size() >= 2) {
//////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////////                        String key = labelCell.getText().trim();
//////////////////                        String value = valueCell.getText().trim();
//////////////////                        configMap.put(key, value);
//////////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////        } catch (IOException e) {
//////////////////            e.printStackTrace();
//////////////////        }
//////////////////    }
//////////////////
//////////////////    private static String convertDocToDocx(String sourceFile) {
//////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////////////        try {
//////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////////////            doc.save(docxFile, SaveFormat.DOCX);
//////////////////        } catch (Exception e) {
//////////////////            throw new RuntimeException(e);
//////////////////        }
//////////////////        return docxFile;
//////////////////    }
//////////////////
//////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////////
//////////////////            StringBuilder originalContent = new StringBuilder();
//////////////////            StringBuilder modifiedContent = new StringBuilder();
//////////////////
//////////////////            List<XWPFTable> tables = document.getTables();
//////////////////
//////////////////            for (XWPFTable table : tables) {
//////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////////////                        originalContent.append(text).append(" ");
//////////////////
//////////////////                        if (configMap.containsKey(text)) {
//////////////////                            String newValue = configMap.get(text);
//////////////////                            modifiedContent.append(newValue).append(" ");
//////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////////                                nextCell.removeParagraph(0);
//////////////////                                XWPFParagraph p = nextCell.addParagraph();
//////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////////////                                XWPFRun r = p.createRun();
//////////////////                                r.setText(newValue);
//////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////////////                            }
//////////////////                        }
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////
//////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////////                List<XWPFRun> runs = paragraph.getRuns();
//////////////////                for (int i = 0; i < runs.size(); i++) {
//////////////////                    XWPFRun run = runs.get(i);
//////////////////                    String text = run.getText(0);
//////////////////                    if (text != null) {
//////////////////                        originalContent.append(text).append(" ");
//////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////////                            String key = entry.getKey();
//////////////////                            String value = entry.getValue();
//////////////////                            if (text.trim().equals(key)) {
//////////////////                                int j = i + 1;
//////////////////                                while (j < runs.size()) {
//////////////////                                    XWPFRun nextRun = runs.get(j);
//////////////////                                    String nextText = nextRun.getText(0);
//////////////////                                    if (nextText != null && !nextText.contains(":")) {
//////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////////////                                        newRun.setText(value);
//////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////////////                                        newRun.setFontSize(14);
//////////////////                                        paragraph.removeRun(j);
//////////////////                                        break;
//////////////////                                    }
//////////////////                                    j++;
//////////////////                                }
//////////////////                                i = j;
//////////////////                                break;
//////////////////                            }
//////////////////                        }
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////
//////////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////////////            document.write(fos);
//////////////////        }
//////////////////    }
//////////////////
//////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////////////        progressBar.setValue(progress);
//////////////////    }
//////////////////
//////////////////    private static JFrame createMainFrame() {
//////////////////        JFrame frame = new JFrame("文档处理工具");
//////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////////////        frame.setSize(1200, 800);
//////////////////
//////////////////        JPanel panel = new JPanel(new BorderLayout());
//////////////////        frame.add(panel);
//////////////////
//////////////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////////////        configTable = new JTable(configTableModel);
//////////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////////////
//////////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////////////        configInputPanel.add(new JLabel("Key:"));
//////////////////        configInputPanel.add(keyField);
//////////////////        configInputPanel.add(new JLabel("Value:"));
//////////////////        configInputPanel.add(valueField);
//////////////////
//////////////////        JButton addButton = new JButton("添加/更新");
//////////////////        addButton.addActionListener(new ActionListener() {
//////////////////            @Override
//////////////////            public void actionPerformed(ActionEvent e) {
//////////////////                String key = keyField.getText().trim();
//////////////////                String value = valueField.getText().trim();
//////////////////                if (!key.isEmpty() && !value.isEmpty()) {
//////////////////                    configMap.put(key, value);
//////////////////                    boolean keyExists = false;
//////////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////////////                            configTableModel.setValueAt(value, i, 1);
//////////////////                            keyExists = true;
//////////////////                            break;
//////////////////                        }
//////////////////                    }
//////////////////                    if (!keyExists) {
//////////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////        });
//////////////////        configInputPanel.add(addButton);
//////////////////
//////////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////////////
//////////////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////////////////        fileTable = new JTable(fileTableModel);
//////////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
//////////////////            if (!event.getValueIsAdjusting()) {
//////////////////                int selectedRow = fileTable.getSelectedRow();
//////////////////                if (selectedRow >= 0) {
//////////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
//////////////////                }
//////////////////            }
//////////////////        });
//////////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////////////
//////////////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////////////
//////////////////        JPanel statsPanel = new JPanel(new BorderLayout());
//////////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("文档处理统计分析"));
//////////////////        statsTextArea.setEditable(false);
//////////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////////////////
//////////////////        JButton startButton = new JButton("开始执行");
//////////////////        startButton.addActionListener(new ActionListener() {
//////////////////            @Override
//////////////////            public void actionPerformed(ActionEvent e) {
//////////////////                processFiles();
//////////////////            }
//////////////////        });
//////////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////////////////
//////////////////        panel.add(configPanel, BorderLayout.NORTH);
//////////////////        panel.add(filePanel, BorderLayout.CENTER);
//////////////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////////////        panel.add(statsPanel, BorderLayout.EAST);
//////////////////
//////////////////        return frame;
//////////////////    }
//////////////////
//////////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////////////////        previewFrame.setSize(600, 400);
//////////////////        JTextArea previewTextArea = new JTextArea();
//////////////////        previewTextArea.setEditable(false);
//////////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////////////////        previewFrame.add(new JScrollPane(previewTextArea));
//////////////////        previewFrame.setVisible(true);
//////////////////    }
//////////////////
//////////////////    private static void processFiles() {
//////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////////////        File[] listOfFiles = folder.listFiles();
//////////////////        if (listOfFiles == null) {
//////////////////            return;
//////////////////        }
//////////////////        int totalFiles = listOfFiles.length;
//////////////////        int processedFiles = 0;
//////////////////        for (File file : listOfFiles) {
//////////////////            if (file.isFile()) {
//////////////////                String sourceFile = file.getAbsolutePath();
//////////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////////////                if (sourceFile.endsWith(".doc")) {
//////////////////                    sourceFile = convertDocToDocx(sourceFile);
//////////////////                }
//////////////////                if (!sourceFile.endsWith(".docx")) {
//////////////////                    continue;
//////////////////                }
//////////////////                try {
//////////////////                    modifyDocument(sourceFile, outputFile);
//////////////////                    processedFiles++;
//////////////////                    updateProgress(processedFiles, totalFiles);
//////////////////                } catch (IOException e) {
//////////////////                    e.printStackTrace();
//////////////////                }
//////////////////            }
//////////////////        }
//////////////////        displayStats(totalFiles, processedFiles);
//////////////////    }
//////////////////
//////////////////    private static void displayStats(int totalFiles, int processedFiles) {
//////////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n";
//////////////////        statsTextArea.setText(stats);
//////////////////    }
//////////////////}
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////package org.example;
////////////////
////////////////import com.aspose.words.SaveFormat;
////////////////import org.apache.poi.xwpf.usermodel.*;
////////////////
////////////////import javax.swing.*;
////////////////import javax.swing.table.DefaultTableModel;
////////////////import java.awt.*;
////////////////import java.awt.event.ActionEvent;
////////////////import java.awt.event.ActionListener;
////////////////import java.io.File;
////////////////import java.io.FileInputStream;
////////////////import java.io.FileOutputStream;
////////////////import java.io.IOException;
////////////////import java.util.HashMap;
////////////////import java.util.List;
////////////////import java.util.Map;
////////////////
////////////////public class WordModifier {
////////////////    private static Map<String, String> configMap = new HashMap<>();
////////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////////    private static JFrame frame;
////////////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////////////////    private static JTextField valueField = new JTextField();
////////////////    private static JTable configTable;
////////////////    private static JTable fileTable;
////////////////    private static JTextArea statsTextArea = new JTextArea();
////////////////
////////////////    public static void main(String[] args) {
////////////////        // 创建并显示主窗口
////////////////        frame = createMainFrame();
////////////////        frame.setVisible(true);
////////////////
////////////////        // 加载配置文件
////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////////        loadConfigFile(configFile);
////////////////    }
////////////////
////////////////    private static void loadConfigFile(String configFile) {
////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////////
////////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////////
////////////////            for (XWPFTable table : configTables) {
////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////                    if (row.getTableCells().size() >= 2) {
////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////////                        String key = labelCell.getText().trim();
////////////////                        String value = valueCell.getText().trim();
////////////////                        configMap.put(key, value);
////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////                        keyComboBox.addItem(key);
////////////////                    }
////////////////                }
////////////////            }
////////////////        } catch (IOException e) {
////////////////            e.printStackTrace();
////////////////        }
////////////////    }
////////////////
////////////////    private static void saveConfigFile(String configFile) {
////////////////        try (FileOutputStream fos = new FileOutputStream(configFile);
////////////////             XWPFDocument configDoc = new XWPFDocument()) {
////////////////
////////////////            XWPFTable table = configDoc.createTable();
////////////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////                XWPFTableRow row = table.createRow();
////////////////                row.getCell(0).setText(entry.getKey());
////////////////                row.getCell(1).setText(entry.getValue());
////////////////            }
////////////////            configDoc.write(fos);
////////////////        } catch (IOException e) {
////////////////            e.printStackTrace();
////////////////        }
////////////////    }
////////////////
////////////////    private static String convertDocToDocx(String sourceFile) {
////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////////        try {
////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////////        } catch (Exception e) {
////////////////            throw new RuntimeException(e);
////////////////        }
////////////////        return docxFile;
////////////////    }
////////////////
////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////////
////////////////            StringBuilder originalContent = new StringBuilder();
////////////////            StringBuilder modifiedContent = new StringBuilder();
////////////////
////////////////            List<XWPFTable> tables = document.getTables();
////////////////
////////////////            for (XWPFTable table : tables) {
////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////////                        originalContent.append(text).append(" ");
////////////////
////////////////                        if (configMap.containsKey(text)) {
////////////////                            String newValue = configMap.get(text);
////////////////                            modifiedContent.append(newValue).append(" ");
////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////////                                nextCell.removeParagraph(0);
////////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////////                                XWPFRun r = p.createRun();
////////////////                                r.setText(newValue);
////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////////                            }
////////////////                        }
////////////////                    }
////////////////                }
////////////////            }
////////////////
////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////////                for (int i = 0; i < runs.size(); i++) {
////////////////                    XWPFRun run = runs.get(i);
////////////////                    String text = run.getText(0);
////////////////                    if (text != null) {
////////////////                        originalContent.append(text).append(" ");
////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////                            String key = entry.getKey();
////////////////                            String value = entry.getValue();
////////////////                            if (text.trim().equals(key)) {
////////////////                                int j = i + 1;
////////////////                                while (j < runs.size()) {
////////////////                                    XWPFRun nextRun = runs.get(j);
////////////////                                    String nextText = nextRun.getText(0);
////////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////////                                        newRun.setText(value);
////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////////                                        newRun.setFontSize(14);
////////////////                                        paragraph.removeRun(j);
////////////////                                        break;
////////////////                                    }
////////////////                                    j++;
////////////////                                }
////////////////                                i = j;
////////////////                                break;
////////////////                            }
////////////////                        }
////////////////                    }
////////////////                }
////////////////            }
////////////////
////////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////////////            document.write(fos);
////////////////        }
////////////////    }
////////////////
////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////////        progressBar.setValue(progress);
////////////////    }
////////////////
////////////////    private static JFrame createMainFrame() {
////////////////        JFrame frame = new JFrame("文档处理工具");
////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////////        frame.setSize(1200, 800);
////////////////
////////////////        JPanel panel = new JPanel(new BorderLayout());
////////////////        frame.add(panel);
////////////////
////////////////        JPanel configPanel = new JPanel(new BorderLayout());
////////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////////////        configTable = new JTable(configTableModel);
////////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////////////
////////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////////////        configInputPanel.add(new JLabel("Key:"));
////////////////        configInputPanel.add(keyComboBox);
////////////////        configInputPanel.add(new JLabel("Value:"));
////////////////        configInputPanel.add(valueField);
////////////////
////////////////        JButton addButton = new JButton("添加/更新");
////////////////        addButton.addActionListener(new ActionListener() {
////////////////            @Override
////////////////            public void actionPerformed(ActionEvent e) {
////////////////                String key = (String) keyComboBox.getSelectedItem();
////////////////                String value = valueField.getText().trim();
////////////////                if (!key.isEmpty() && !value.isEmpty()) {
////////////////                    configMap.put(key, value);
////////////////                    boolean keyExists = false;
////////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////////////                            configTableModel.setValueAt(value, i, 1);
////////////////                            keyExists = true;
////////////////                            break;
////////////////                        }
////////////////                    }
////////////////                    if (!keyExists) {
////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////                        keyComboBox.addItem(key);
////////////////                    }
////////////////                    saveConfigFile("Z:\\Desktop\\测试\\模板\\模板.docx");
////////////////                }
////////////////            }
////////////////        });
////////////////        configInputPanel.add(addButton);
////////////////
////////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////////////
////////////////        JPanel filePanel = new JPanel(new BorderLayout());
////////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////////////        fileTable = new JTable(fileTableModel);
////////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
////////////////            if (!event.getValueIsAdjusting()) {
////////////////                int selectedRow = fileTable.getSelectedRow();
////////////////                if (selectedRow != -1) {
////////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
////////////////                }
////////////////            }
////////////////        });
////////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////////////
////////////////        JButton refreshButton = new JButton("刷新预览");
////////////////        refreshButton.addActionListener(new ActionListener() {
////////////////            @Override
////////////////            public void actionPerformed(ActionEvent e) {
////////////////                fileTableModel.setRowCount(0);
////////////////                processFiles();
////////////////            }
////////////////        });
////////////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
////////////////
////////////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////////////
////////////////        JButton startButton = new JButton("开始执行");
////////////////        startButton.addActionListener(new ActionListener() {
////////////////            @Override
////////////////            public void actionPerformed(ActionEvent e) {
////////////////                processFiles();
////////////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
////////////////            }
////////////////        });
////////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////////////
////////////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
////////////////        statsTextArea.setEditable(false);
////////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////////////
////////////////        panel.add(configPanel, BorderLayout.NORTH);
////////////////        panel.add(filePanel, BorderLayout.CENTER);
////////////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////////////        panel.add(statsPanel, BorderLayout.EAST);
////////////////
////////////////        return frame;
////////////////    }
////////////////
////////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////////////        previewFrame.setSize(600, 400);
////////////////        JTextArea previewTextArea = new JTextArea();
////////////////        previewTextArea.setEditable(false);
////////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////////////        previewFrame.add(new JScrollPane(previewTextArea));
////////////////        previewFrame.setVisible(true);
////////////////    }
////////////////
////////////////    private static void processFiles() {
////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////////        File[] listOfFiles = folder.listFiles();
////////////////        if (listOfFiles == null) {
////////////////            return;
////////////////        }
////////////////        int totalFiles = listOfFiles.length;
////////////////        int processedFiles = 0;
////////////////        long startTime = System.currentTimeMillis();
////////////////        for (File file : listOfFiles) {
////////////////            if (file.isFile()) {
////////////////                String sourceFile = file.getAbsolutePath();
////////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////////                if (sourceFile.endsWith(".doc")) {
////////////////                    sourceFile = convertDocToDocx(sourceFile);
////////////////                }
////////////////                if (!sourceFile.endsWith(".docx")) {
////////////////                    continue;
////////////////                }
////////////////                try {
////////////////                    modifyDocument(sourceFile, outputFile);
////////////////                    processedFiles++;
////////////////                    updateProgress(processedFiles, totalFiles);
////////////////                } catch (IOException e) {
////////////////                    e.printStackTrace();
////////////////                }
////////////////            }
////////////////        }
////////////////        long endTime = System.currentTimeMillis();
////////////////        long duration = endTime - startTime;
////////////////        displayStats(totalFiles, processedFiles, duration);
////////////////    }
////////////////
////////////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
////////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
////////////////        statsTextArea.setText(stats);
////////////////    }
////////////////}
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////package org.example;
//////////////
//////////////import com.aspose.words.SaveFormat;
//////////////import org.apache.poi.xwpf.usermodel.*;
//////////////
//////////////import javax.swing.*;
//////////////import javax.swing.table.DefaultTableModel;
//////////////import java.awt.*;
//////////////import java.awt.event.ActionEvent;
//////////////import java.awt.event.ActionListener;
//////////////import java.io.File;
//////////////import java.io.FileInputStream;
//////////////import java.io.FileOutputStream;
//////////////import java.io.IOException;
//////////////import java.util.HashMap;
//////////////import java.util.List;
//////////////import java.util.Map;
//////////////
//////////////public class WordModifier {
//////////////    private static Map<String, String> configMap = new HashMap<>();
//////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////////    private static JFrame frame;
//////////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////////////    private static JTextField valueField = new JTextField();
//////////////    private static JTable configTable;
//////////////    private static JTable fileTable;
//////////////    private static JTextArea statsTextArea = new JTextArea();
//////////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////////
//////////////    public static void main(String[] args) {
//////////////        // 创建并显示主窗口
//////////////        frame = createMainFrame();
//////////////        frame.setVisible(true);
//////////////
//////////////        // 加载配置文件
//////////////        loadConfigFile(CONFIG_FILE);
//////////////    }
//////////////
//////////////    private static void loadConfigFile(String configFile) {
//////////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////
//////////////            List<XWPFTable> configTables = configDoc.getTables();
//////////////
//////////////            for (XWPFTable table : configTables) {
//////////////                for (XWPFTableRow row : table.getRows()) {
//////////////                    if (row.getTableCells().size() >= 2) {
//////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////                        String key = labelCell.getText().trim();
//////////////                        String value = valueCell.getText().trim();
//////////////                        configMap.put(key, value);
//////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////                        keyComboBox.addItem(key);
//////////////                    }
//////////////                }
//////////////            }
//////////////        } catch (IOException e) {
//////////////            e.printStackTrace();
//////////////        }
//////////////    }
//////////////
//////////////    private static void saveConfigFile(String configFile) {
//////////////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////////////             XWPFDocument configDoc = new XWPFDocument()) {
//////////////
//////////////            XWPFTable table = configDoc.createTable();
//////////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////                XWPFTableRow row = table.createRow();
//////////////                row.getCell(0).setText(entry.getKey());
//////////////                row.getCell(1).setText(entry.getValue());
//////////////            }
//////////////            configDoc.write(fos);
//////////////        } catch (IOException e) {
//////////////            e.printStackTrace();
//////////////        }
//////////////    }
//////////////
//////////////    private static String convertDocToDocx(String sourceFile) {
//////////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////////        try {
//////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////////            doc.save(docxFile, SaveFormat.DOCX);
//////////////        } catch (Exception e) {
//////////////            throw new RuntimeException(e);
//////////////        }
//////////////        return docxFile;
//////////////    }
//////////////
//////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////
//////////////            StringBuilder originalContent = new StringBuilder();
//////////////            StringBuilder modifiedContent = new StringBuilder();
//////////////
//////////////            List<XWPFTable> tables = document.getTables();
//////////////
//////////////            for (XWPFTable table : tables) {
//////////////                for (XWPFTableRow row : table.getRows()) {
//////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////////                        originalContent.append(text).append(" ");
//////////////
//////////////                        if (configMap.containsKey(text)) {
//////////////                            String newValue = configMap.get(text);
//////////////                            modifiedContent.append(newValue).append(" ");
//////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////                                nextCell.removeParagraph(0);
//////////////                                XWPFParagraph p = nextCell.addParagraph();
//////////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////////                                XWPFRun r = p.createRun();
//////////////                                r.setText(newValue);
//////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////////                            }
//////////////                        }
//////////////                    }
//////////////                }
//////////////            }
//////////////
//////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////                List<XWPFRun> runs = paragraph.getRuns();
//////////////                for (int i = 0; i < runs.size(); i++) {
//////////////                    XWPFRun run = runs.get(i);
//////////////                    String text = run.getText(0);
//////////////                    if (text != null) {
//////////////                        originalContent.append(text).append(" ");
//////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////                            String key = entry.getKey();
//////////////                            String value = entry.getValue();
//////////////                            if (text.trim().equals(key)) {
//////////////                                int j = i + 1;
//////////////                                while (j < runs.size()) {
//////////////                                    XWPFRun nextRun = runs.get(j);
//////////////                                    String nextText = nextRun.getText(0);
//////////////                                    if (nextText != null && !nextText.contains(":")) {
//////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////////                                        newRun.setText(value);
//////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////////                                        newRun.setFontSize(14);
//////////////                                        paragraph.removeRun(j);
//////////////                                        break;
//////////////                                    }
//////////////                                    j++;
//////////////                                }
//////////////                                i = j;
//////////////                                break;
//////////////                            }
//////////////                        }
//////////////                    }
//////////////                }
//////////////            }
//////////////
//////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////////            document.write(fos);
//////////////        }
//////////////    }
//////////////
//////////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////////        progressBar.setValue(progress);
//////////////    }
//////////////
//////////////    private static JFrame createMainFrame() {
//////////////        JFrame frame = new JFrame("文档处理工具");
//////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////////        frame.setSize(1200, 800);
//////////////
//////////////        JPanel panel = new JPanel(new BorderLayout());
//////////////        frame.add(panel);
//////////////
//////////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////////        configTable = new JTable(configTableModel);
//////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////////
//////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////////        configInputPanel.add(new JLabel("Key:"));
//////////////        configInputPanel.add(keyComboBox);
//////////////        configInputPanel.add(new JLabel("Value:"));
//////////////        configInputPanel.add(valueField);
//////////////
//////////////        JButton addButton = new JButton("添加/更新");
//////////////        addButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                String key = (String) keyComboBox.getSelectedItem();
//////////////                String value = valueField.getText().trim();
//////////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
//////////////                    configMap.put(key, value);
//////////////                    boolean keyExists = false;
//////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////////                            configTableModel.setValueAt(value, i, 1);
//////////////                            keyExists = true;
//////////////                            break;
//////////////                        }
//////////////                    }
//////////////                    if (!keyExists) {
//////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////                        keyComboBox.addItem(key);
//////////////                    }
//////////////                    saveConfigFile(CONFIG_FILE);
//////////////                }
//////////////            }
//////////////        });
//////////////        configInputPanel.add(addButton);
//////////////
//////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////////
//////////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////////////        fileTable = new JTable(fileTableModel);
//////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
//////////////            if (!event.getValueIsAdjusting()) {
//////////////                int selectedRow = fileTable.getSelectedRow();
//////////////                if (selectedRow != -1) {
//////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
//////////////                }
//////////////            }
//////////////        });
//////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////////
//////////////        JButton refreshButton = new JButton("刷新预览");
//////////////        refreshButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                fileTableModel.setRowCount(0);
//////////////                processFiles();
//////////////            }
//////////////        });
//////////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////////////
//////////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////////
//////////////        JButton startButton = new JButton("开始执行");
//////////////        startButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                processFiles();
//////////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////////////            }
//////////////        });
//////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////////////
//////////////        JPanel statsPanel = new JPanel(new BorderLayout());
//////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////////////        statsTextArea.setEditable(false);
//////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////////////
//////////////        panel.add(configPanel, BorderLayout.NORTH);
//////////////        panel.add(filePanel, BorderLayout.CENTER);
//////////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////////        panel.add(statsPanel, BorderLayout.EAST);
//////////////
//////////////        return frame;
//////////////    }
//////////////
//////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////////////        previewFrame.setSize(600, 400);
//////////////        JTextArea previewTextArea = new JTextArea();
//////////////        previewTextArea.setEditable(false);
//////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////////////        previewFrame.add(new JScrollPane(previewTextArea));
//////////////        previewFrame.setVisible(true);
//////////////    }
//////////////
//////////////    private static void processFiles() {
//////////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////////        File[] listOfFiles = folder.listFiles();
//////////////        if (listOfFiles == null) {
//////////////            return;
//////////////        }
//////////////        int totalFiles = listOfFiles.length;
//////////////        int processedFiles = 0;
//////////////        long startTime = System.currentTimeMillis();
//////////////        for (File file : listOfFiles) {
//////////////            if (file.isFile()) {
//////////////                String sourceFile = file.getAbsolutePath();
//////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////////                if (sourceFile.endsWith(".doc")) {
//////////////                    sourceFile = convertDocToDocx(sourceFile);
//////////////                }
//////////////                if (!sourceFile.endsWith(".docx")) {
//////////////                    continue;
//////////////                }
//////////////                try {
//////////////                    modifyDocument(sourceFile, outputFile);
//////////////                    processedFiles++;
//////////////                    updateProgress(processedFiles, totalFiles);
//////////////                } catch (IOException e) {
//////////////                    e.printStackTrace();
//////////////                }
//////////////            }
//////////////        }
//////////////        long endTime = System.currentTimeMillis();
//////////////        long duration = endTime - startTime;
//////////////        displayStats(totalFiles, processedFiles, duration);
//////////////    }
//////////////
//////////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////////////        statsTextArea.setText(stats);
//////////////    }
//////////////}
////////////
////////////
////////////
////////////package org.example;
////////////
////////////import com.aspose.words.SaveFormat;
////////////import org.apache.poi.xwpf.usermodel.*;
////////////
////////////import javax.swing.*;
////////////import javax.swing.table.DefaultTableModel;
////////////import java.awt.*;
////////////import java.awt.event.ActionEvent;
////////////import java.awt.event.ActionListener;
////////////import java.io.File;
////////////import java.io.FileInputStream;
////////////import java.io.FileOutputStream;
////////////import java.io.IOException;
////////////import java.util.HashMap;
////////////import java.util.List;
////////////import java.util.Map;
////////////
////////////public class WordModifier {
////////////    private static Map<String, String> configMap = new HashMap<>();
////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////    private static JFrame frame;
////////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////////////    private static JTextField valueField = new JTextField();
////////////    private static JTable configTable;
////////////    private static JTable fileTable;
////////////    private static JTextArea statsTextArea = new JTextArea();
////////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////
////////////    public static void main(String[] args) {
////////////        // 创建并显示主窗口
////////////        frame = createMainFrame();
////////////        frame.setVisible(true);
////////////
////////////        // 加载配置文件
////////////        loadConfigFile(CONFIG_FILE);
////////////    }
////////////
////////////    private static void loadConfigFile(String configFile) {
////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////
////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////
////////////            for (XWPFTable table : configTables) {
////////////                for (XWPFTableRow row : table.getRows()) {
////////////                    if (row.getTableCells().size() >= 2) {
////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////                        String key = labelCell.getText().trim();
////////////                        String value = valueCell.getText().trim();
////////////                        configMap.put(key, value);
////////////                        configTableModel.addRow(new Object[]{key, value});
////////////                        keyComboBox.addItem(key);
////////////                    }
////////////                }
////////////            }
////////////        } catch (IOException e) {
////////////            e.printStackTrace();
////////////        }
////////////    }
////////////
////////////    private static void saveConfigFile(String configFile) {
////////////        try (FileOutputStream fos = new FileOutputStream(configFile);
////////////             XWPFDocument configDoc = new XWPFDocument()) {
////////////
////////////            XWPFTable table = configDoc.createTable();
////////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////                XWPFTableRow row = table.createRow();
////////////                row.getCell(0).setText(entry.getKey());
////////////                row.getCell(1).setText(entry.getValue());
////////////            }
////////////            configDoc.write(fos);
////////////        } catch (IOException e) {
////////////            e.printStackTrace();
////////////        }
////////////    }
////////////
////////////    private static String convertDocToDocx(String sourceFile) {
////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////        try {
////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////        } catch (Exception e) {
////////////            throw new RuntimeException(e);
////////////        }
////////////        return docxFile;
////////////    }
////////////
////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////
////////////            StringBuilder originalContent = new StringBuilder();
////////////            StringBuilder modifiedContent = new StringBuilder();
////////////
////////////            List<XWPFTable> tables = document.getTables();
////////////
////////////            for (XWPFTable table : tables) {
////////////                for (XWPFTableRow row : table.getRows()) {
////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////                        originalContent.append(text).append(" ");
////////////
////////////                        if (configMap.containsKey(text)) {
////////////                            String newValue = configMap.get(text);
////////////                            modifiedContent.append(newValue).append(" ");
////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////                                nextCell.removeParagraph(0);
////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////                                XWPFRun r = p.createRun();
////////////                                r.setText(newValue);
////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////                            }
////////////                        }
////////////                    }
////////////                }
////////////            }
////////////
////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////                for (int i = 0; i < runs.size(); i++) {
////////////                    XWPFRun run = runs.get(i);
////////////                    String text = run.getText(0);
////////////                    if (text != null) {
////////////                        originalContent.append(text).append(" ");
////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////                            String key = entry.getKey();
////////////                            String value = entry.getValue();
////////////                            if (text.trim().equals(key)) {
////////////                                int j = i + 1;
////////////                                while (j < runs.size()) {
////////////                                    XWPFRun nextRun = runs.get(j);
////////////                                    String nextText = nextRun.getText(0);
////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////                                        newRun.setText(value);
////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////                                        newRun.setFontSize(14);
////////////                                        paragraph.removeRun(j);
////////////                                        break;
////////////                                    }
////////////                                    j++;
////////////                                }
////////////                                i = j;
////////////                                break;
////////////                            }
////////////                        }
////////////                    }
////////////                }
////////////            }
////////////
////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////////            document.write(fos);
////////////        }
////////////    }
////////////
////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////        progressBar.setValue(progress);
////////////    }
////////////
////////////    private static JFrame createMainFrame() {
////////////        JFrame frame = new JFrame("文档处理工具");
////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////        frame.setSize(1200, 800);
////////////
////////////        JPanel panel = new JPanel(new BorderLayout());
////////////        frame.add(panel);
////////////
////////////        JPanel configPanel = new JPanel(new BorderLayout());
////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////////        configTable = new JTable(configTableModel);
////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////////
////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////////        configInputPanel.add(new JLabel("Key:"));
////////////        configInputPanel.add(keyComboBox);
////////////        configInputPanel.add(new JLabel("Value:"));
////////////        configInputPanel.add(valueField);
////////////
////////////        JButton addButton = new JButton("添加/更新");
////////////        addButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                String key = (String) keyComboBox.getSelectedItem();
////////////                String value = valueField.getText().trim();
////////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
////////////                    configMap.put(key, value);
////////////                    boolean keyExists = false;
////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////////                            configTableModel.setValueAt(value, i, 1);
////////////                            keyExists = true;
////////////                            break;
////////////                        }
////////////                    }
////////////                    if (!keyExists) {
////////////                        configTableModel.addRow(new Object[]{key, value});
////////////                        keyComboBox.addItem(key);
////////////                    }
////////////                    saveConfigFile(CONFIG_FILE);
////////////                }
////////////            }
////////////        });
////////////        configInputPanel.add(addButton);
////////////
////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////////
////////////        JPanel filePanel = new JPanel(new BorderLayout());
////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////////        fileTable = new JTable(fileTableModel);
////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
////////////            if (!event.getValueIsAdjusting()) {
////////////                int selectedRow = fileTable.getSelectedRow();
////////////                if (selectedRow != -1) {
////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
////////////                }
////////////            }
////////////        });
////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////////
////////////        JButton refreshButton = new JButton("刷新预览");
////////////        refreshButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                fileTableModel.setRowCount(0);
////////////                processFiles();
////////////            }
////////////        });
////////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
////////////
////////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////////
////////////        JButton startButton = new JButton("开始执行");
////////////        startButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                processFiles();
////////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
////////////            }
////////////        });
////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////////
////////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
////////////        statsTextArea.setEditable(false);
////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////////
////////////        panel.add(configPanel, BorderLayout.NORTH);
////////////        panel.add(filePanel, BorderLayout.CENTER);
////////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////////        panel.add(statsPanel, BorderLayout.EAST);
////////////
////////////        return frame;
////////////    }
////////////
////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////////        previewFrame.setSize(600, 400);
////////////        JTextArea previewTextArea = new JTextArea();
////////////        previewTextArea.setEditable(false);
////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////////        previewFrame.add(new JScrollPane(previewTextArea));
////////////        previewFrame.setVisible(true);
////////////    }
////////////
////////////    private static void processFiles() {
////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////        File[] listOfFiles = folder.listFiles();
////////////        if (listOfFiles == null) {
////////////            return;
////////////        }
////////////        int totalFiles = listOfFiles.length;
////////////        int processedFiles = 0;
////////////        long startTime = System.currentTimeMillis();
////////////        for (File file : listOfFiles) {
////////////            if (file.isFile()) {
////////////                String sourceFile = file.getAbsolutePath();
////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////                if (sourceFile.endsWith(".doc")) {
////////////                    sourceFile = convertDocToDocx(sourceFile);
////////////                }
////////////                if (!sourceFile.endsWith(".docx")) {
////////////                    continue;
////////////                }
////////////                try {
////////////                    modifyDocument(sourceFile, outputFile);
////////////                    processedFiles++;
////////////                    updateProgress(processedFiles, totalFiles);
////////////                } catch (IOException e) {
////////////                    e.printStackTrace();
////////////                }
////////////            }
////////////        }
////////////        long endTime = System.currentTimeMillis();
////////////        long duration = endTime - startTime;
////////////        displayStats(totalFiles, processedFiles, duration);
////////////    }
////////////
////////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
////////////        statsTextArea.setText(stats);
////////////    }
////////////}
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////
//////////package org.example;
//////////
//////////import com.aspose.words.SaveFormat;
//////////import org.apache.poi.xwpf.usermodel.*;
//////////
//////////import javax.swing.*;
//////////import javax.swing.table.DefaultTableModel;
//////////import java.awt.*;
//////////import java.awt.event.ActionEvent;
//////////import java.awt.event.ActionListener;
//////////import java.io.File;
//////////import java.io.FileInputStream;
//////////import java.io.FileOutputStream;
//////////import java.io.IOException;
//////////import java.util.HashMap;
//////////import java.util.List;
//////////import java.util.Map;
//////////
//////////public class WordModifier {
//////////    private static Map<String, String> configMap = new HashMap<>();
//////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////    private static JFrame frame;
//////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////////    private static JTextField valueField = new JTextField();
//////////    private static JTable configTable;
//////////    private static JTable fileTable;
//////////    private static JTextArea statsTextArea = new JTextArea();
//////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////
//////////    public static void main(String[] args) {
//////////        // 创建并显示主窗口
//////////        frame = createMainFrame();
//////////        frame.setVisible(true);
//////////
//////////        // 加载配置文件
//////////        loadConfigFile(CONFIG_FILE);
//////////    }
//////////
//////////    private static void loadConfigFile(String configFile) {
//////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////
//////////            List<XWPFTable> configTables = configDoc.getTables();
//////////
//////////            for (XWPFTable table : configTables) {
//////////                for (XWPFTableRow row : table.getRows()) {
//////////                    if (row.getTableCells().size() >= 2) {
//////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////                        String key = labelCell.getText().trim();
//////////                        String value = valueCell.getText().trim();
//////////                        configMap.put(key, value);
//////////                        configTableModel.addRow(new Object[]{key, value});
//////////                        keyComboBox.addItem(key);
//////////                    }
//////////                }
//////////            }
//////////        } catch (IOException e) {
//////////            e.printStackTrace();
//////////        }
//////////    }
//////////
//////////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//////////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////////             XWPFDocument configDoc = new XWPFDocument()) {
//////////
//////////            XWPFTable table = configDoc.createTable();
//////////            boolean firstRow = true;
//////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////                XWPFTableRow row;
//////////                if (firstRow) {
//////////                    row = table.getRow(0); // 使用已经存在的第一行
//////////                    firstRow = false;
//////////                } else {
//////////                    row = table.createRow();
//////////                }
//////////                row.getCell(0).setText(entry.getKey());
//////////                row.getCell(1).setText(entry.getValue());
//////////            }
//////////            configDoc.write(fos);
//////////        } catch (IOException e) {
//////////            // 这里可以使用更合适的异常处理方式
//////////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//////////        }
//////////    }
//////////
//////////
//////////    private static String convertDocToDocx(String sourceFile) {
//////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////        try {
//////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////            doc.save(docxFile, SaveFormat.DOCX);
//////////        } catch (Exception e) {
//////////            throw new RuntimeException(e);
//////////        }
//////////        return docxFile;
//////////    }
//////////
//////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////
//////////            StringBuilder originalContent = new StringBuilder();
//////////            StringBuilder modifiedContent = new StringBuilder();
//////////
//////////            List<XWPFTable> tables = document.getTables();
//////////
//////////            for (XWPFTable table : tables) {
//////////                for (XWPFTableRow row : table.getRows()) {
//////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////                        originalContent.append(text).append(" ");
//////////
//////////                        if (configMap.containsKey(text)) {
//////////                            String newValue = configMap.get(text);
//////////                            modifiedContent.append(newValue).append(" ");
//////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////                                nextCell.removeParagraph(0);
//////////                                XWPFParagraph p = nextCell.addParagraph();
//////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////                                XWPFRun r = p.createRun();
//////////                                r.setText(newValue);
//////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////                            }
//////////                        }
//////////                    }
//////////                }
//////////            }
//////////
//////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////                List<XWPFRun> runs = paragraph.getRuns();
//////////                for (int i = 0; i < runs.size(); i++) {
//////////                    XWPFRun run = runs.get(i);
//////////                    String text = run.getText(0);
//////////                    if (text != null) {
//////////                        originalContent.append(text).append(" ");
//////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////                            String key = entry.getKey();
//////////                            String value = entry.getValue();
//////////                            if (text.trim().equals(key)) {
//////////                                int j = i + 1;
//////////                                while (j < runs.size()) {
//////////                                    XWPFRun nextRun = runs.get(j);
//////////                                    String nextText = nextRun.getText(0);
//////////                                    if (nextText != null && !nextText.contains(":")) {
//////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////                                        newRun.setText(value);
//////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////                                        newRun.setFontSize(14);
//////////                                        paragraph.removeRun(j);
//////////                                        break;
//////////                                    }
//////////                                    j++;
//////////                                }
//////////                                i = j;
//////////                                break;
//////////                            }
//////////                        }
//////////                    }
//////////                }
//////////            }
//////////
//////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////            document.write(fos);
//////////        }
//////////    }
//////////
//////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////        progressBar.setValue(progress);
//////////    }
//////////
//////////    private static JFrame createMainFrame() {
//////////        JFrame frame = new JFrame("文档处理工具");
//////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////        frame.setSize(1200, 800);
//////////
//////////        JPanel panel = new JPanel(new BorderLayout());
//////////        frame.add(panel);
//////////
//////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////        configTable = new JTable(configTableModel);
//////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////
//////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////        configInputPanel.add(new JLabel("Key:"));
//////////        configInputPanel.add(keyComboBox);
//////////        configInputPanel.add(new JLabel("Value:"));
//////////        configInputPanel.add(valueField);
//////////
//////////        JButton addButton = new JButton("添加/更新");
//////////        addButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                String key = (String) keyComboBox.getSelectedItem();
//////////                String value = valueField.getText().trim();
//////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
//////////                    configMap.put(key, value);
//////////                    boolean keyExists = false;
//////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////                            configTableModel.setValueAt(value, i, 1);
//////////                            keyExists = true;
//////////                            break;
//////////                        }
//////////                    }
//////////                    if (!keyExists) {
//////////                        configTableModel.addRow(new Object[]{key, value});
//////////                        keyComboBox.addItem(key);
//////////                    }
//////////                    saveConfigFile(CONFIG_FILE,configMap);
//////////                }
//////////            }
//////////        });
//////////        configInputPanel.add(addButton);
//////////
//////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////
//////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////////        fileTable = new JTable(fileTableModel);
//////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
//////////            if (!event.getValueIsAdjusting()) {
//////////                int selectedRow = fileTable.getSelectedRow();
//////////                if (selectedRow != -1) {
//////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////////                    displayFilePreview(fileName, originalContent, modifiedContent);
//////////                }
//////////            }
//////////        });
//////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////
//////////        JButton refreshButton = new JButton("刷新预览");
//////////        refreshButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                fileTableModel.setRowCount(0);
//////////                processFiles();
//////////            }
//////////        });
//////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////////
//////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////
//////////        JButton startButton = new JButton("开始执行");
//////////        startButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                processFiles();
//////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////////            }
//////////        });
//////////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////////
//////////        JPanel statsPanel = new JPanel(new BorderLayout());
//////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////////        statsTextArea.setEditable(false);
//////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////////
//////////        panel.add(configPanel, BorderLayout.NORTH);
//////////        panel.add(filePanel, BorderLayout.CENTER);
//////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////        panel.add(statsPanel, BorderLayout.EAST);
//////////
//////////        return frame;
//////////    }
//////////
//////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////////        previewFrame.setSize(600, 400);
//////////        JTextArea previewTextArea = new JTextArea();
//////////        previewTextArea.setEditable(false);
//////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////////        previewFrame.add(new JScrollPane(previewTextArea));
//////////        previewFrame.setVisible(true);
//////////    }
//////////
//////////    private static void processFiles() {
//////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////        File[] listOfFiles = folder.listFiles();
//////////        if (listOfFiles == null) {
//////////            return;
//////////        }
//////////        int totalFiles = listOfFiles.length;
//////////        int processedFiles = 0;
//////////        long startTime = System.currentTimeMillis();
//////////        for (File file : listOfFiles) {
//////////            if (file.isFile()) {
//////////                String sourceFile = file.getAbsolutePath();
//////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////                if (sourceFile.endsWith(".doc")) {
//////////                    sourceFile = convertDocToDocx(sourceFile);
//////////                }
//////////                if (!sourceFile.endsWith(".docx")) {
//////////                    continue;
//////////                }
//////////                try {
//////////                    modifyDocument(sourceFile, outputFile);
//////////                    processedFiles++;
//////////                    updateProgress(processedFiles, totalFiles);
//////////                } catch (IOException e) {
//////////                    e.printStackTrace();
//////////                }
//////////            }
//////////        }
//////////        long endTime = System.currentTimeMillis();
//////////        long duration = endTime - startTime;
//////////        displayStats(totalFiles, processedFiles, duration);
//////////    }
//////////
//////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////////        statsTextArea.setText(stats);
//////////    }
//////////}
//////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////package org.example;
////////
////////import com.aspose.words.SaveFormat;
////////import org.apache.poi.xwpf.usermodel.*;
////////
////////import javax.swing.*;
////////import javax.swing.event.ListSelectionEvent;
////////import javax.swing.event.ListSelectionListener;
////////import javax.swing.table.DefaultTableModel;
////////import java.awt.*;
////////import java.awt.event.ActionEvent;
////////import java.awt.event.ActionListener;
////////import java.io.File;
////////import java.io.FileInputStream;
////////import java.io.FileOutputStream;
////////import java.io.IOException;
////////import java.util.HashMap;
////////import java.util.LinkedHashMap;
////////import java.util.List;
////////import java.util.Map;
////////
////////public class WordModifier {
////////    private static Map<String, String> configMap = new LinkedHashMap<>();
////////
////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////    private static JFrame frame;
////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////////    private static JTextField valueField = new JTextField();
////////    private static JTable configTable;
////////    private static JTable fileTable;
////////    private static JTextArea statsTextArea = new JTextArea();
////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////
////////    public static void main(String[] args) {
////////        // 创建并显示主窗口
////////        frame = createMainFrame();
////////        frame.setVisible(true);
////////
////////        // 加载配置文件
////////        loadConfigFile(CONFIG_FILE);
////////    }
////////
////////    private static void loadConfigFile(String configFile) {
////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////
////////            List<XWPFTable> configTables = configDoc.getTables();
////////
////////            for (XWPFTable table : configTables) {
////////                for (XWPFTableRow row : table.getRows()) {
////////                    if (row.getTableCells().size() >= 2) {
////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////                        String key = labelCell.getText().trim();
////////                        String value = valueCell.getText().trim();
////////                        configMap.put(key, value);
////////                        configTableModel.addRow(new Object[]{key, value});
////////                        keyComboBox.addItem(key);
////////                    }
////////                }
////////            }
////////        } catch (IOException e) {
////////            e.printStackTrace();
////////        }
////////    }
////////
////////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
////////        try (FileOutputStream fos = new FileOutputStream(configFile);
////////             XWPFDocument configDoc = new XWPFDocument()) {
////////
////////            XWPFTable table = configDoc.createTable();
////////            boolean firstRow = true;
////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////                XWPFTableRow row;
////////                if (firstRow) {
////////                    row = table.getRow(0); // 使用已经存在的第一行
////////                    firstRow = false;
////////                } else {
////////                    row = table.createRow();
////////                }
////////                XWPFTableCell keyCell = row.getCell(0);
////////                if (keyCell == null) {
////////                    keyCell = row.addNewTableCell();
////////                }
////////                keyCell.setText(entry.getKey());
////////
////////                XWPFTableCell valueCell;
////////                if (row.getTableCells().size() > 1) {
////////                    valueCell = row.getCell(1);
////////                } else {
////////                    valueCell = row.addNewTableCell();
////////                }
////////                valueCell.setText(entry.getValue());
////////            }
////////            configDoc.write(fos);
////////        } catch (IOException e) {
////////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
////////        }
////////    }
////////
////////
////////    private static String convertDocToDocx(String sourceFile) {
////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////        try {
////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////            doc.save(docxFile, SaveFormat.DOCX);
////////        } catch (Exception e) {
////////            throw new RuntimeException(e);
////////        }
////////        return docxFile;
////////    }
////////
////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////
////////            StringBuilder originalContent = new StringBuilder();
////////            StringBuilder modifiedContent = new StringBuilder();
////////
////////            List<XWPFTable> tables = document.getTables();
////////
////////            for (XWPFTable table : tables) {
////////                for (XWPFTableRow row : table.getRows()) {
////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////                        String text = cell.getText().replaceAll("\\s+", "");
////////                        originalContent.append(text).append(" ");
////////
////////                        if (configMap.containsKey(text)) {
////////                            String newValue = configMap.get(text);
////////                            modifiedContent.append(newValue).append(" ");
////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////                                nextCell.removeParagraph(0);
////////                                XWPFParagraph p = nextCell.addParagraph();
////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////                                XWPFRun r = p.createRun();
////////                                r.setText(newValue);
////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////                            }
////////                        }
////////                    }
////////                }
////////            }
////////
////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////                List<XWPFRun> runs = paragraph.getRuns();
////////                for (int i = 0; i < runs.size(); i++) {
////////                    XWPFRun run = runs.get(i);
////////                    String text = run.getText(0);
////////                    if (text != null) {
////////                        originalContent.append(text).append(" ");
////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////                            String key = entry.getKey();
////////                            String value = entry.getValue();
////////                            if (text.trim().equals(key)) {
////////                                int j = i + 1;
////////                                while (j < runs.size()) {
////////                                    XWPFRun nextRun = runs.get(j);
////////                                    String nextText = nextRun.getText(0);
////////                                    if (nextText != null && !nextText.contains(":")) {
////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////                                        newRun.setText(value);
////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////                                        newRun.setFontFamily("仿宋_GB2312");
////////                                        newRun.setFontSize(14);
////////                                        paragraph.removeRun(j);
////////                                        break;
////////                                    }
////////                                    j++;
////////                                }
////////                                i = j;
////////                                break;
////////                            }
////////                        }
////////                    }
////////                }
////////            }
////////
////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////            document.write(fos);
////////        }
////////    }
////////
////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////        progressBar.setValue(progress);
////////    }
////////
////////    private static JFrame createMainFrame() {
////////        JFrame frame = new JFrame("文档处理工具");
////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////        frame.setSize(1200, 800);
////////
////////        JPanel panel = new JPanel(new BorderLayout());
////////        frame.add(panel);
////////
////////        JPanel configPanel = new JPanel(new BorderLayout());
////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////        configTable = new JTable(configTableModel);
////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////
////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////        configInputPanel.add(new JLabel("Key:"));
////////        configInputPanel.add(keyComboBox);
////////        configInputPanel.add(new JLabel("Value:"));
////////        configInputPanel.add(valueField);
////////
////////        JButton addButton = new JButton("添加/更新");
////////        addButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                String key = (String) keyComboBox.getSelectedItem();
////////                String value = valueField.getText().trim();
////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
////////                    configMap.put(key, value);
////////                    boolean keyExists = false;
////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////                            configTableModel.setValueAt(value, i, 1);
////////                            keyExists = true;
////////                            break;
////////                        }
////////                    }
////////                    if (!keyExists) {
////////                        configTableModel.addRow(new Object[]{key, value});
////////                        keyComboBox.addItem(key);
////////                    }
////////                    saveConfigFile(CONFIG_FILE, configMap);
////////                }
////////            }
////////        });
////////        configInputPanel.add(addButton);
////////
////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////
////////        JPanel filePanel = new JPanel(new BorderLayout());
////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////        fileTable = new JTable(fileTableModel);
////////        fileTable.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
////////            @Override
////////            public void valueChanged(ListSelectionEvent event) {
////////                if (!event.getValueIsAdjusting()) {
////////                    int selectedRow = fileTable.getSelectedRow();
////////                    if (selectedRow != -1) {
////////                        String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////                        String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////                        String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////                        displayFilePreview(fileName, originalContent, modifiedContent);
////////                    }
////////                }
////////            }
////////        });
////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////
////////        JButton refreshButton = new JButton("刷新预览");
////////        refreshButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                fileTableModel.setRowCount(0);
////////                processFiles();
////////            }
////////        });
////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
////////
////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////
////////        JButton startButton = new JButton("开始执行");
////////        startButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                processFiles();
////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
////////            }
////////        });
////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////
////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
////////        statsTextArea.setEditable(false);
////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////
////////        panel.add(configPanel, BorderLayout.NORTH);
////////        panel.add(filePanel, BorderLayout.CENTER);
////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////        panel.add(statsPanel, BorderLayout.EAST);
////////
////////        return frame;
////////    }
////////
////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////        previewFrame.setSize(600, 400);
////////        JTextArea previewTextArea = new JTextArea();
////////        previewTextArea.setEditable(false);
////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////        previewFrame.add(new JScrollPane(previewTextArea));
////////        previewFrame.setVisible(true);
////////    }
////////
////////    private static void processFiles() {
////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////        File[] listOfFiles = folder.listFiles();
////////        if (listOfFiles == null) {
////////            return;
////////        }
////////        int totalFiles = listOfFiles.length;
////////        int processedFiles = 0;
////////        long startTime = System.currentTimeMillis();
////////        for (File file : listOfFiles) {
////////            if (file.isFile()) {
////////                String sourceFile = file.getAbsolutePath();
////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////                if (sourceFile.endsWith(".doc")) {
////////                    sourceFile = convertDocToDocx(sourceFile);
////////                }
////////                if (!sourceFile.endsWith(".docx")) {
////////                    continue;
////////                }
////////                try {
////////                    modifyDocument(sourceFile, outputFile);
////////                    processedFiles++;
////////                    updateProgress(processedFiles, totalFiles);
////////                } catch (IOException e) {
////////                    e.printStackTrace();
////////                }
////////            }
////////        }
////////        long endTime = System.currentTimeMillis();
////////        long duration = endTime - startTime;
////////        displayStats(totalFiles, processedFiles, duration);
////////    }
////////
////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
////////        statsTextArea.setText(stats);
////////    }
////////}
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////package org.example;
//////
//////import com.aspose.words.SaveFormat;
//////import org.apache.poi.xwpf.usermodel.*;
//////
//////import javax.swing.*;
//////import javax.swing.event.ListSelectionEvent;
//////import javax.swing.event.ListSelectionListener;
//////import javax.swing.table.DefaultTableModel;
//////import java.awt.*;
//////import java.awt.event.ActionEvent;
//////import java.awt.event.ActionListener;
//////import java.io.File;
//////import java.io.FileInputStream;
//////import java.io.FileOutputStream;
//////import java.io.IOException;
//////import java.util.*;
//////import java.util.List;
//////
//////public class WordModifier {
//////    private static Map<String, String> configMap = new LinkedHashMap<>();
//////
//////    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
//////    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();
//////
//////
//////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////    private static JFrame frame;
//////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////    private static JTextField valueField = new JTextField();
//////    private static JTable configTable;
//////    private static JTable fileTable;
//////    private static JTextArea statsTextArea = new JTextArea();
//////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////
//////    public static void main(String[] args) {
//////        // 创建并显示主窗口
//////        frame = createMainFrame();
//////        frame.setVisible(true);
//////
//////        // 加载配置文件
//////        loadConfigFile(CONFIG_FILE);
//////    }
//////
//////    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
//////        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
//////             XWPFDocument doc = new XWPFDocument(fis)) {
//////
//////            List<XWPFTable> tables = doc.getTables();
//////            for (XWPFTable table : tables) {
//////                for (XWPFTableRow row : table.getRows()) {
//////                    List<XWPFTableCell> cells = row.getTableCells();
//////                    if (cells.size() >= 5) {
//////                        String alias = cells.get(0).getText().trim();
//////                        List<String> keys = new ArrayList<>();
//////                        for (int i = 1; i < cells.size() - 1; i++) {
//////                            String key = cells.get(i).getText().trim();
//////                            if (!key.isEmpty() && !key.equals("-")) {
//////                                keys.add(key);
//////                            }
//////                        }
//////                        // Skip this row if no keys were found
//////                        if (keys.isEmpty()) {
//////                            continue;
//////                        }
//////                        String value = cells.get(cells.size() - 1).getText().trim();
//////                        aliasToKeysMap.put(alias, keys);
//////                        aliasToValueMap.put(alias, value);
//////                    }
//////                }
//////            }
//////        } catch (IOException e) {
//////            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
//////        }
//////    }
//////
//////
//////    private static void loadConfigFile(String configFile) {
//////        try (FileInputStream configFis = new FileInputStream(configFile);
//////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////
//////            List<XWPFTable> configTables = configDoc.getTables();
//////
//////            for (XWPFTable table : configTables) {
//////                for (XWPFTableRow row : table.getRows()) {
//////                    if (row.getTableCells().size() >= 2) {
//////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////                        String key = labelCell.getText().trim();
//////                        String value = valueCell.getText().trim();
//////                        configMap.put(key, value);
//////                        configTableModel.addRow(new Object[]{key, value});
//////                        keyComboBox.addItem(key);
//////                    }
//////                }
//////            }
//////        } catch (IOException e) {
//////            e.printStackTrace();
//////        }
//////        // After loading the original config file, load the equivalent keys file
//////        loadEquivalentKeysFile("Z:\\Desktop\\测试\\模板\\模板2.docx");
//////
//////        // Merge the equivalent keys into the original config
//////        mergeEquivalentKeysToConfig();
//////    }
//////    private static void mergeEquivalentKeysToConfig() {
//////        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//////            String alias = entry.getKey();
//////            String value = aliasToValueMap.get(alias);
//////            for (String key : entry.getValue()) {
//////                configMap.put(key, value);
//////                updateConfigTableModel(key, value);
//////            }
//////        }
//////    }
//////    private static void updateConfigTableModel(String key, String value) {
//////        boolean keyExists = false;
//////        for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////            if (configTableModel.getValueAt(i, 0).equals(key)) {
//////                configTableModel.setValueAt(value, i, 1);
//////                keyExists = true;
//////                break;
//////            }
//////        }
//////        if (!keyExists) {
//////            configTableModel.addRow(new Object[]{key, value});
//////        }
//////    }
//////
//////
//////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////             XWPFDocument configDoc = new XWPFDocument()) {
//////
//////            XWPFTable table = configDoc.createTable();
//////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////                XWPFTableRow row = table.createRow();
//////                XWPFTableCell keyCell = row.addNewTableCell();
//////                keyCell.setText(entry.getKey());
//////
//////                XWPFTableCell valueCell = row.addNewTableCell();
//////                valueCell.setText(entry.getValue());
//////            }
//////            configDoc.write(fos);
//////        } catch (IOException e) {
//////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//////        }
//////    }
//////
//////
//////    private static String convertDocToDocx(String sourceFile) {
//////        String docxFile = sourceFile.replace(".doc", ".docx");
//////        try {
//////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////            doc.save(docxFile, SaveFormat.DOCX);
//////        } catch (Exception e) {
//////            throw new RuntimeException(e);
//////        }
//////        return docxFile;
//////    }
//////
//////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////             FileOutputStream fos = new FileOutputStream(outputFile);
//////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////
//////            StringBuilder originalContent = new StringBuilder();
//////            StringBuilder modifiedContent = new StringBuilder();
//////
//////            List<XWPFTable> tables = document.getTables();
//////
//////            for (XWPFTable table : tables) {
//////                for (XWPFTableRow row : table.getRows()) {
//////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////                        XWPFTableCell cell = row.getTableCells().get(i);
//////                        String text = cell.getText().replaceAll("\\s+", "");
//////                        originalContent.append(text).append(" ");
//////
//////                        if (configMap.containsKey(text)) {
//////                            String newValue = configMap.get(text);
//////                            modifiedContent.append(newValue).append(" ");
//////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////                                nextCell.removeParagraph(0);
//////                                XWPFParagraph p = nextCell.addParagraph();
//////                                p.setAlignment(ParagraphAlignment.CENTER);
//////                                XWPFRun r = p.createRun();
//////                                r.setText(newValue);
//////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////                            }
//////                        }
//////                    }
//////                }
//////            }
//////
//////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////                List<XWPFRun> runs = paragraph.getRuns();
//////                for (int i = 0; i < runs.size(); i++) {
//////                    XWPFRun run = runs.get(i);
//////                    String text = run.getText(0);
//////                    if (text != null) {
//////                        originalContent.append(text).append(" ");
//////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////                            String key = entry.getKey();
//////                            String value = entry.getValue();
//////                            if (text.trim().equals(key)) {
//////                                int j = i + 1;
//////                                while (j < runs.size()) {
//////                                    XWPFRun nextRun = runs.get(j);
//////                                    String nextText = nextRun.getText(0);
//////                                    if (nextText != null && !nextText.contains(":")) {
//////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////                                        newRun.setText(value);
//////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////                                        newRun.setFontFamily("仿宋_GB2312");
//////                                        newRun.setFontSize(14);
//////                                        paragraph.removeRun(j);
//////                                        break;
//////                                    }
//////                                    j++;
//////                                }
//////                                i = j;
//////                                break;
//////                            }
//////                        }
//////                    }
//////                }
//////            }
//////
//////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////            document.write(fos);
//////        }
//////    }
//////
//////    private static void updateProgress(int processedFiles, int totalFiles) {
//////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////        progressBar.setValue(progress);
//////    }
//////
//////    private static JFrame createMainFrame() {
//////        JFrame frame = new JFrame("文档处理工具");
//////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////        frame.setSize(1200, 800);
//////
//////        JPanel panel = new JPanel(new BorderLayout());
//////        frame.add(panel);
//////
//////        JPanel configPanel = new JPanel(new BorderLayout());
//////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////        configTable = new JTable(configTableModel);
//////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////
//////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////        configInputPanel.add(new JLabel("Key:"));
//////        configInputPanel.add(keyComboBox);
//////        configInputPanel.add(new JLabel("Value:"));
//////        configInputPanel.add(valueField);
//////
//////        JButton addButton = new JButton("添加/更新");
//////        addButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                String key = (String) keyComboBox.getSelectedItem();
//////                String value = valueField.getText().trim();
//////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
//////                    configMap.put(key, value);
//////                    boolean keyExists = false;
//////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////                            configTableModel.setValueAt(value, i, 1);
//////                            keyExists = true;
//////                            break;
//////                        }
//////                    }
//////                    if (!keyExists) {
//////                        configTableModel.addRow(new Object[]{key, value});
//////                        keyComboBox.addItem(key);
//////                    }
//////                    saveConfigFile(CONFIG_FILE, configMap);
//////                }
//////            }
//////        });
//////        configInputPanel.add(addButton);
//////
//////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////
//////        JPanel filePanel = new JPanel(new BorderLayout());
//////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////        fileTable = new JTable(fileTableModel);
//////        fileTable.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
//////            @Override
//////            public void valueChanged(ListSelectionEvent event) {
//////                if (!event.getValueIsAdjusting()) {
//////                    int selectedRow = fileTable.getSelectedRow();
//////                    if (selectedRow != -1) {
//////                        String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////                        String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////                        String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////                        displayFilePreview(fileName, originalContent, modifiedContent);
//////                    }
//////                }
//////            }
//////        });
//////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////
//////        JButton refreshButton = new JButton("刷新预览");
//////        refreshButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                fileTableModel.setRowCount(0);
//////                processFiles();
//////            }
//////        });
//////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////
//////        JPanel progressPanel = new JPanel(new BorderLayout());
//////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////
//////        JButton startButton = new JButton("开始执行");
//////        startButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                processFiles();
//////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////            }
//////        });
//////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////
//////        JPanel statsPanel = new JPanel(new BorderLayout());
//////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////        statsTextArea.setEditable(false);
//////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////
//////        panel.add(configPanel, BorderLayout.NORTH);
//////        panel.add(filePanel, BorderLayout.CENTER);
//////        panel.add(progressPanel, BorderLayout.SOUTH);
//////        panel.add(statsPanel, BorderLayout.EAST);
//////
//////        return frame;
//////    }
//////
//////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////        previewFrame.setSize(600, 400);
//////        JTextArea previewTextArea = new JTextArea();
//////        previewTextArea.setEditable(false);
//////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////        previewFrame.add(new JScrollPane(previewTextArea));
//////        previewFrame.setVisible(true);
//////    }
//////
//////    private static void processFiles() {
//////        File folder = new File("Z:\\Desktop\\测试\\in");
//////        File[] listOfFiles = folder.listFiles();
//////        if (listOfFiles == null) {
//////            return;
//////        }
//////        int totalFiles = listOfFiles.length;
//////        int processedFiles = 0;
//////        long startTime = System.currentTimeMillis();
//////        for (File file : listOfFiles) {
//////            if (file.isFile()) {
//////                String sourceFile = file.getAbsolutePath();
//////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////                if (sourceFile.endsWith(".doc")) {
//////                    sourceFile = convertDocToDocx(sourceFile);
//////                }
//////                if (!sourceFile.endsWith(".docx")) {
//////                    continue;
//////                }
//////                try {
//////                    modifyDocument(sourceFile, outputFile);
//////                    processedFiles++;
//////                    updateProgress(processedFiles, totalFiles);
//////                } catch (IOException e) {
//////                    e.printStackTrace();
//////                }
//////            }
//////        }
//////        long endTime = System.currentTimeMillis();
//////        long duration = endTime - startTime;
//////        displayStats(totalFiles, processedFiles, duration);
//////    }
//////
//////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////        statsTextArea.setText(stats);
//////    }
//////}
////
////
////
////
////
////
////
////
////
////
////
////
////
////
////
////
////package org.example;
////
////import com.aspose.words.SaveFormat;
////import org.apache.poi.xwpf.usermodel.*;
////
////import javax.swing.*;
////import javax.swing.table.DefaultTableModel;
////import java.awt.*;
////import java.awt.event.ActionEvent;
////import java.awt.event.ActionListener;
////import java.io.File;
////import java.io.FileInputStream;
////import java.io.FileOutputStream;
////import java.io.IOException;
////import java.util.*;
////import java.util.List;
////
////public class WordModifier {
////    private static Map<String, String> configMap = new LinkedHashMap<>();
////    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
////    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();
////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////    private static JFrame frame;
////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////    private static JTextField valueField = new JTextField();
////    private static JTable configTable;
////    private static JTable fileTable;
////    private static JTextArea statsTextArea = new JTextArea();
////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
////    private static final String EQUIVALENT_KEYS_FILE = "Z:\\Desktop\\测试\\模板\\模板2.docx";
////
////    public static void main(String[] args) {
////        // 创建并显示主窗口
////        frame = createMainFrame();
////        frame.setVisible(true);
////
////        // 加载配置文件
////        loadConfigFile(CONFIG_FILE);
////
////        // 加载等价键配置文件
////        loadEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
////
////        // 合并等价键到原有配置文件中
////        mergeEquivalentKeysToConfig();
////
////        // 显示等价键设置界面
////        displayEquivalentKeysUI();
////    }
////
////    private static void loadConfigFile(String configFile) {
////        try (FileInputStream configFis = new FileInputStream(configFile);
////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////
////            List<XWPFTable> configTables = configDoc.getTables();
////
////            for (XWPFTable table : configTables) {
////                for (XWPFTableRow row : table.getRows()) {
////                    if (row.getTableCells().size() >= 2) {
////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////                        String key = labelCell.getText().trim();
////                        String value = valueCell.getText().trim();
////                        configMap.put(key, value);
////                        configTableModel.addRow(new Object[]{key, value});
////                        keyComboBox.addItem(key);
////                    }
////                }
////            }
////        } catch (IOException e) {
////            e.printStackTrace();
////        }
////    }
////
////    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
////        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
////             XWPFDocument doc = new XWPFDocument(fis)) {
////
////            List<XWPFTable> tables = doc.getTables();
////            for (XWPFTable table : tables) {
////                for (XWPFTableRow row : table.getRows()) {
////                    List<XWPFTableCell> cells = row.getTableCells();
////                    if (cells.size() >= 5) {
////                        String alias = cells.get(0).getText().trim();
////                        List<String> keys = new ArrayList<>();
////                        for (int i = 1; i < cells.size() - 1; i++) {
////                            String key = cells.get(i).getText().trim();
////                            if (!key.isEmpty() && !key.equals("-")) {
////                                keys.add(key);
////                            }
////                        }
////                        String value = cells.get(cells.size() - 1).getText().trim();
////                        aliasToKeysMap.put(alias, keys);
////                        aliasToValueMap.put(alias, value);
////                    }
////                }
////            }
////        } catch (IOException e) {
////            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
////        }
////    }
////
////    private static void mergeEquivalentKeysToConfig() {
////        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
////            String alias = entry.getKey();
////            String value = aliasToValueMap.get(alias);
////            for (String key : entry.getValue()) {
////                configMap.put(key, value);
////                updateConfigTableModel(key, value);
////            }
////        }
////    }
////
////    private static void updateConfigTableModel(String key, String value) {
////        boolean keyExists = false;
////        for (int i = 0; i < configTableModel.getRowCount(); i++) {
////            if (configTableModel.getValueAt(i, 0).equals(key)) {
////                configTableModel.setValueAt(value, i, 1);
////                keyExists = true;
////                break;
////            }
////        }
////        if (!keyExists) {
////            configTableModel.addRow(new Object[]{key, value});
////            keyComboBox.addItem(key);
////        }
////    }
////
////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
////        try (FileOutputStream fos = new FileOutputStream(configFile);
////             XWPFDocument configDoc = new XWPFDocument()) {
////
////            XWPFTable table = configDoc.createTable();
////            boolean firstRow = true;
////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////                XWPFTableRow row;
////                if (firstRow) {
////                    row = table.getRow(0); // 使用已经存在的第一行
////                    firstRow = false;
////                } else {
////                    row = table.createRow();
////                }
////                row.getCell(0).setText(entry.getKey());
////                row.addNewTableCell().setText(entry.getValue()); // addNewTableCell to create a new cell
////            }
////            configDoc.write(fos);
////        } catch (IOException e) {
////            // 这里可以使用更合适的异常处理方式
////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
////        }
////    }
////
////    private static String convertDocToDocx(String sourceFile) {
////        String docxFile = sourceFile.replace(".doc", ".docx");
////        try {
////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////            doc.save(docxFile, SaveFormat.DOCX);
////        } catch (Exception e) {
////            throw new RuntimeException(e);
////        }
////        return docxFile;
////    }
////
////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////             FileOutputStream fos = new FileOutputStream(outputFile);
////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////
////            StringBuilder originalContent = new StringBuilder();
////            StringBuilder modifiedContent = new StringBuilder();
////
////            List<XWPFTable> tables = document.getTables();
////
////            for (XWPFTable table : tables) {
////                for (XWPFTableRow row : table.getRows()) {
////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////                        XWPFTableCell cell = row.getTableCells().get(i);
////                        String text = cell.getText().replaceAll("\\s+", "");
////                        originalContent.append(text).append(" ");
////
////                        if (configMap.containsKey(text)) {
////                            String newValue = configMap.get(text);
////                            modifiedContent.append(newValue).append(" ");
////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////                                nextCell.removeParagraph(0);
////                                XWPFParagraph p = nextCell.addParagraph();
////                                p.setAlignment(ParagraphAlignment.CENTER);
////                                XWPFRun r = p.createRun();
////                                r.setText(newValue);
////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////                            }
////                        }
////                    }
////                }
////            }
////
////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////                List<XWPFRun> runs = paragraph.getRuns();
////                for (int i = 0; i < runs.size(); i++) {
////                    XWPFRun run = runs.get(i);
////                    String text = run.getText(0);
////                    if (text != null) {
////                        originalContent.append(text).append(" ");
////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////                            String key = entry.getKey();
////                            String value = entry.getValue();
////                            if (text.trim().equals(key)) {
////                                int j = i + 1;
////                                while (j < runs.size()) {
////                                    XWPFRun nextRun = runs.get(j);
////                                    String nextText = nextRun.getText(0);
////                                    if (nextText != null && !nextText.contains(":")) {
////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////                                        newRun.setText(value);
////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////                                        newRun.setFontFamily("仿宋");
////                                        newRun.setFontSize(12);
////                                        break;
////                                    }
////                                    j++;
////                                }
////                            }
////                        }
////                    }
////                }
////            }
////
////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////
////            document.write(fos);
////        }
////    }
////
////    private static void displayEquivalentKeysUI() {
////        JFrame aliasFrame = new JFrame("设置等价键值");
////        aliasFrame.setSize(400, 300);
////        JPanel aliasPanel = new JPanel(new BorderLayout());
////
////        DefaultTableModel aliasTableModel = new DefaultTableModel(new Object[]{"别名", "键", "值"}, 0);
////        JTable aliasTable = new JTable(aliasTableModel);
////
////        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
////            String alias = entry.getKey();
////            String value = aliasToValueMap.get(alias);
////            for (String key : entry.getValue()) {
////                aliasTableModel.addRow(new Object[]{alias, key, value});
////            }
////        }
////
////        aliasPanel.add(new JScrollPane(aliasTable), BorderLayout.CENTER);
////
////        JButton saveButton = new JButton("保存");
////        saveButton.addActionListener(e -> {
////            // 处理保存逻辑
////            // 你可以添加代码来保存用户在界面上修改的值
////        });
////        aliasPanel.add(saveButton, BorderLayout.SOUTH);
////
////        aliasFrame.add(aliasPanel);
////        aliasFrame.setVisible(true);
////    }
////
////    private static JFrame createMainFrame() {
////        JFrame frame = new JFrame("Word Modifier");
////        frame.setSize(800, 600);
////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////        frame.setLayout(new BorderLayout());
////
////        JTabbedPane tabbedPane = new JTabbedPane();
////
////        // Configuration Panel
////        JPanel configPanel = new JPanel(new BorderLayout());
////        configTable = new JTable(configTableModel);
////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////        tabbedPane.addTab("Configuration", configPanel);
////
////        // File Panel
////        JPanel filePanel = new JPanel(new BorderLayout());
////        fileTable = new JTable(fileTableModel);
////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////        tabbedPane.addTab("Files", filePanel);
////
////        frame.add(tabbedPane, BorderLayout.CENTER);
////
////        // Control Panel
////        JPanel controlPanel = new JPanel();
////        controlPanel.add(new JLabel("Key:"));
////        controlPanel.add(keyComboBox);
////        controlPanel.add(new JLabel("Value:"));
////        valueField.setPreferredSize(new Dimension(100, 24));
////        controlPanel.add(valueField);
////        JButton updateButton = new JButton("Update");
////        updateButton.addActionListener(new ActionListener() {
////            @Override
////            public void actionPerformed(ActionEvent e) {
////                String key = (String) keyComboBox.getSelectedItem();
////                String value = valueField.getText();
////                if (key != null && !value.isEmpty()) {
////                    configMap.put(key, value);
////                    updateConfigTableModel(key, value);
////                    saveConfigFile(CONFIG_FILE, configMap);
////                }
////            }
////        });
////        controlPanel.add(updateButton);
////
////        JButton processFilesButton = new JButton("Process Files");
////        processFilesButton.addActionListener(new ActionListener() {
////            @Override
////            public void actionPerformed(ActionEvent e) {
////                processFiles();
////            }
////        });
////        controlPanel.add(processFilesButton);
////        frame.add(controlPanel, BorderLayout.SOUTH);
////
////        frame.add(progressBar, BorderLayout.NORTH);
////
////        // Statistics Panel
////        JPanel statsPanel = new JPanel(new BorderLayout());
////        statsTextArea.setEditable(false);
////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////        tabbedPane.addTab("Statistics", statsPanel);
////
////        return frame;
////    }
////
////    private static void processFiles() {
////        JFileChooser fileChooser = new JFileChooser();
////        fileChooser.setMultiSelectionEnabled(true);
////        int returnValue = fileChooser.showOpenDialog(frame);
////        if (returnValue == JFileChooser.APPROVE_OPTION) {
////            File[] selectedFiles = fileChooser.getSelectedFiles();
////            progressBar.setMaximum(selectedFiles.length);
////            progressBar.setValue(0);
////            int processedFiles = 0;
////            for (File file : selectedFiles) {
////                String sourceFile = file.getAbsolutePath();
////                String outputFile = sourceFile.replace(".doc", "_modified.docx").replace(".docx", "_modified.docx");
////
////                // Convert .doc files to .docx
////                if (sourceFile.toLowerCase().endsWith(".doc")) {
////                    sourceFile = convertDocToDocx(sourceFile);
////                }
////
////                try {
////                    modifyDocument(sourceFile, outputFile);
////                    processedFiles++;
////                    progressBar.setValue(processedFiles);
////                } catch (IOException e) {
////                    e.printStackTrace();
////                }
////            }
////            showStats();
////        }
////    }
////
////    private static void showStats() {
////        StringBuilder stats = new StringBuilder();
////        stats.append("处理的文件数: ").append(fileTableModel.getRowCount()).append("\n");
////        for (int i = 0; i < fileTableModel.getRowCount(); i++) {
////            String fileName = (String) fileTableModel.getValueAt(i, 0);
////            String originalContent = (String) fileTableModel.getValueAt(i, 1);
////            String modifiedContent = (String) fileTableModel.getValueAt(i, 2);
////            stats.append("文件: ").append(fileName).append("\n")
////                    .append("原始内容: ").append(originalContent).append("\n")
////                    .append("修改后内容: ").append(modifiedContent).append("\n");
////        }
////        statsTextArea.setText(stats.toString());
////    }
////}
//
//
//
//
//
//
//
//package org.example;
//
//import com.aspose.words.SaveFormat;
//import org.apache.poi.xwpf.usermodel.*;
//
//import javax.swing.*;
//import javax.swing.table.DefaultTableModel;
//import java.awt.*;
//import java.awt.event.ActionEvent;
//import java.awt.event.ActionListener;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.*;
//import java.util.List;
//
//public class WordModifier {
//    private static Map<String, String> configMap = new LinkedHashMap<>();
//    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
//    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();
//    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//    private static JProgressBar progressBar = new JProgressBar(0, 100);
//    private static JFrame frame;
//    private static JComboBox<String> keyComboBox = new JComboBox<>();
//    private static JTextField valueField = new JTextField();
//    private static JTable configTable;
//    private static JTable fileTable;
//    private static JTextArea statsTextArea = new JTextArea();
//    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//    private static final String EQUIVALENT_KEYS_FILE = "Z:\\Desktop\\测试\\模板\\模板2.docx";
//
//    public static void main(String[] args) {
//        // 创建并显示主窗口
//        frame = createMainFrame();
//        frame.setVisible(true);
//
//        // 加载配置文件
//        loadConfigFile(CONFIG_FILE);
//
//        // 加载等价键配置文件
//        loadEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
//
//        // 合并等价键到原有配置文件中
//        mergeEquivalentKeysToConfig();
//
//        // 显示等价键设置界面
//        displayEquivalentKeysUI();
//    }
//
//    private static void loadConfigFile(String configFile) {
//        try (FileInputStream configFis = new FileInputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//
//            List<XWPFTable> configTables = configDoc.getTables();
//
//            for (XWPFTable table : configTables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    if (row.getTableCells().size() >= 2) {
//                        XWPFTableCell labelCell = row.getTableCells().get(0);
//                        XWPFTableCell valueCell = row.getTableCells().get(1);
//                        String key = labelCell.getText().trim();
//                        String value = valueCell.getText().trim();
//                        configMap.put(key, value);
//                        configTableModel.addRow(new Object[]{key, value});
//                        keyComboBox.addItem(key);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
//        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
//             XWPFDocument doc = new XWPFDocument(fis)) {
//
//            List<XWPFTable> tables = doc.getTables();
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    List<XWPFTableCell> cells = row.getTableCells();
//                    if (cells.size() >= 5) {
//                        String alias = cells.get(0).getText().trim();
//                        List<String> keys = new ArrayList<>();
//                        for (int i = 1; i < cells.size() - 1; i++) {
//                            String key = cells.get(i).getText().trim();
//                            if (!key.isEmpty() && !key.equals("-")) {
//                                keys.add(key);
//                            }
//                        }
//                        String value = cells.get(cells.size() - 1).getText().trim();
//                        aliasToKeysMap.put(alias, keys);
//                        aliasToValueMap.put(alias, value);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
//        }
//    }
//
//    private static void mergeEquivalentKeysToConfig() {
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                configMap.put(key, value);
//                updateConfigTableModel(key, value);
//            }
//        }
//    }
//
//    private static void updateConfigTableModel(String key, String value) {
//        boolean keyExists = false;
//        for (int i = 0; i < configTableModel.getRowCount(); i++) {
//            if (configTableModel.getValueAt(i, 0).equals(key)) {
//                configTableModel.setValueAt(value, i, 1);
//                keyExists = true;
//                break;
//            }
//        }
//        if (!keyExists) {
//            configTableModel.addRow(new Object[]{key, value});
//            keyComboBox.addItem(key);
//        }
//    }
//
//    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//        try (FileOutputStream fos = new FileOutputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument()) {
//
//            XWPFTable table = configDoc.createTable();
//            boolean firstRow = true;
//            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                XWPFTableRow row;
//                if (firstRow) {
//                    row = table.getRow(0); // 使用已经存在的第一行
//                    firstRow = false;
//                } else {
//                    row = table.createRow();
//                }
//                row.getCell(0).setText(entry.getKey());
//                row.addNewTableCell().setText(entry.getValue()); // addNewTableCell to create a new cell
//            }
//            configDoc.write(fos);
//        } catch (IOException e) {
//            // 这里可以使用更合适的异常处理方式
//            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//        }
//    }
//
//    private static String convertDocToDocx(String sourceFile) {
//        String docxFile = sourceFile.replace(".doc", ".docx");
//        try {
//            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//            doc.save(docxFile, SaveFormat.DOCX);
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
//        return docxFile;
//    }
//
//    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//             FileOutputStream fos = new FileOutputStream(outputFile);
//             XWPFDocument document = new XWPFDocument(sourceFis)) {
//
//            StringBuilder originalContent = new StringBuilder();
//            StringBuilder modifiedContent = new StringBuilder();
//
//            List<XWPFTable> tables = document.getTables();
//
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    for (int i = 0; i < row.getTableCells().size(); i++) {
//                        XWPFTableCell cell = row.getTableCells().get(i);
//                        String text = cell.getText().replaceAll("\\s+", "");
//                        originalContent.append(text).append(" ");
//
//                        if (configMap.containsKey(text)) {
//                            String newValue = configMap.get(text);
//                            modifiedContent.append(newValue).append(" ");
//                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//                                nextCell.removeParagraph(0);
//                                XWPFParagraph p = nextCell.addParagraph();
//                                p.setAlignment(ParagraphAlignment.CENTER);
//                                XWPFRun r = p.createRun();
//                                r.setText(newValue);
//                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                            }
//                        }
//                    }
//                }
//            }
//
//            for (XWPFParagraph paragraph : document.getParagraphs()) {
//                List<XWPFRun> runs = paragraph.getRuns();
//                for (int i = 0; i < runs.size(); i++) {
//                    XWPFRun run = runs.get(i);
//                    String text = run.getText(0);
//                    if (text != null) {
//                        originalContent.append(text).append(" ");
//                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                            String key = entry.getKey();
//                            String value = entry.getValue();
//                            if (text.trim().equals(key)) {
//                                int j = i + 1;
//                                while (j < runs.size()) {
//                                    XWPFRun nextRun = runs.get(j);
//                                    String nextText = nextRun.getText(0);
//                                    if (nextText != null && nextText.trim().equals(value)) {
//                                        run.setText(value, 0);
//                                        modifiedContent.append(value).append(" ");
//                                        nextRun.setText("", 0);
//                                        break;
//                                    }
//                                    j++;
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//
//            document.write(fos);
//        }
//    }
//
//    private static void displayEquivalentKeysUI() {
//        JFrame aliasFrame = new JFrame("设置等价键值");
//        aliasFrame.setSize(400, 300);
//        JPanel aliasPanel = new JPanel(new BorderLayout());
//
//        DefaultTableModel aliasTableModel = new DefaultTableModel(new Object[]{"别名", "键", "值"}, 0);
//        JTable aliasTable = new JTable(aliasTableModel);
//
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                aliasTableModel.addRow(new Object[]{alias, key, value});
//            }
//        }
//
//        aliasPanel.add(new JScrollPane(aliasTable), BorderLayout.CENTER);
//
//        JButton saveButton = new JButton("保存");
//        saveButton.addActionListener(e -> {
//            // 处理保存逻辑
//            aliasToKeysMap.clear();
//            aliasToValueMap.clear();
//            for (int i = 0; i < aliasTableModel.getRowCount(); i++) {
//                String alias = (String) aliasTableModel.getValueAt(i, 0);
//                String key = (String) aliasTableModel.getValueAt(i, 1);
//                String value = (String) aliasTableModel.getValueAt(i, 2);
//
//                aliasToKeysMap.computeIfAbsent(alias, k -> new ArrayList<>()).add(key);
//                aliasToValueMap.put(alias, value);
//            }
//
//            saveEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
//            mergeEquivalentKeysToConfig();
//        });
//        aliasPanel.add(saveButton, BorderLayout.SOUTH);
//
//        aliasFrame.add(aliasPanel);
//        aliasFrame.setVisible(true);
//    }
//
//    private static void saveEquivalentKeysFile(String equivalentKeysFile) {
//        try (FileOutputStream fos = new FileOutputStream(equivalentKeysFile);
//             XWPFDocument doc = new XWPFDocument()) {
//
//            XWPFTable table = doc.createTable();
//            for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//                XWPFTableRow row = table.createRow();
//                row.getCell(0).setText(entry.getKey());
//                int cellIndex = 1;
//                for (String key : entry.getValue()) {
//                    if (cellIndex >= row.getTableCells().size()) {
//                        row.addNewTableCell().setText(key);
//                    } else {
//                        row.getCell(cellIndex).setText(key);
//                    }
//                    cellIndex++;
//                }
//                if (cellIndex >= row.getTableCells().size()) {
//                    row.addNewTableCell().setText(aliasToValueMap.get(entry.getKey()));
//                } else {
//                    row.getCell(cellIndex).setText(aliasToValueMap.get(entry.getKey()));
//                }
//            }
//
//            doc.write(fos);
//        } catch (IOException e) {
//            System.err.println("An error occurred while saving the equivalent keys file: " + e.getMessage());
//        }
//    }
//
//    private static JFrame createMainFrame() {
//        JFrame frame = new JFrame("Word Modifier");
//        frame.setSize(800, 600);
//        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        frame.setLayout(new BorderLayout());
//
//        JTabbedPane tabbedPane = new JTabbedPane();
//
//        // Configuration Panel
//        JPanel configPanel = new JPanel(new BorderLayout());
//        configTable = new JTable(configTableModel);
//        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Configuration", configPanel);
//
//        // File Panel
//        JPanel filePanel = new JPanel(new BorderLayout());
//        fileTable = new JTable(fileTableModel);
//        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Files", filePanel);
//
//        frame.add(tabbedPane, BorderLayout.CENTER);
//
//        // Control Panel
//        JPanel controlPanel = new JPanel();
//        controlPanel.add(new JLabel("Key:"));
//        controlPanel.add(keyComboBox);
//        controlPanel.add(new JLabel("Value:"));
//        valueField.setPreferredSize(new Dimension(100, 24));
//        controlPanel.add(valueField);
//        JButton updateButton = new JButton("Update");
//        updateButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                String key = (String) keyComboBox.getSelectedItem();
//                String value = valueField.getText();
//                if (key != null && !value.isEmpty()) {
//                    configMap.put(key, value);
//                    updateConfigTableModel(key, value);
//                    saveConfigFile(CONFIG_FILE, configMap);
//                }
//            }
//        });
//        controlPanel.add(updateButton);
//
//        JButton processFilesButton = new JButton("Process Files");
//        processFilesButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                processFiles();
//            }
//        });
//        controlPanel.add(processFilesButton);
//        frame.add(controlPanel, BorderLayout.SOUTH);
//
//        frame.add(progressBar, BorderLayout.NORTH);
//
//        // Statistics Panel
//        JPanel statsPanel = new JPanel(new BorderLayout());
//        statsTextArea.setEditable(false);
//        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//        tabbedPane.addTab("Statistics", statsPanel);
//
//        return frame;
//    }
//
//    private static void processFiles() {
//        JFileChooser fileChooser = new JFileChooser();
//        fileChooser.setMultiSelectionEnabled(true);
//        int returnValue = fileChooser.showOpenDialog(frame);
//        if (returnValue == JFileChooser.APPROVE_OPTION) {
//            File[] selectedFiles = fileChooser.getSelectedFiles();
//            progressBar.setMaximum(selectedFiles.length);
//            progressBar.setValue(0);
//            int processedFiles = 0;
//            for (File file : selectedFiles) {
//                String sourceFile = file.getAbsolutePath();
//                String outputFile = sourceFile.replace(".doc", "_modified.docx").replace(".docx", "_modified.docx");
//
//                // Convert .doc files to .docx
//                if (sourceFile.toLowerCase().endsWith(".doc")) {
//                    sourceFile = convertDocToDocx(sourceFile);
//                }
//
//                try {
//                    modifyDocument(sourceFile, outputFile);
//                    processedFiles++;
//                    progressBar.setValue(processedFiles);
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
//            showStats();
//        }
//    }
//
//    private static void showStats() {
//        StringBuilder stats = new StringBuilder();
//        stats.append("处理的文件数: ").append(fileTableModel.getRowCount()).append("\n");
//        for (int i = 0; i < fileTableModel.getRowCount(); i++) {
//            String fileName = (String) fileTableModel.getValueAt(i, 0);
//            String originalContent = (String) fileTableModel.getValueAt(i, 1);
//            String modifiedContent = (String) fileTableModel.getValueAt(i, 2);
//            stats.append("文件: ").append(fileName).append("\n")
//                    .append("原始内容: ").append(originalContent).append("\n")
//                    .append("修改后内容: ").append(modifiedContent).append("\n");
//        }
//        statsTextArea.setText(stats.toString());
//    }
//}





//////////////////////package org.example;
//////////////////////
//////////////////////import com.aspose.words.SaveFormat;
//////////////////////import org.apache.poi.xwpf.usermodel.*;
//////////////////////
//////////////////////import java.io.File;
//////////////////////import java.io.FileInputStream;
//////////////////////import java.io.FileOutputStream;
//////////////////////import java.io.IOException;
//////////////////////import java.util.HashMap;
//////////////////////import java.util.List;
//////////////////////import java.util.Map;
//////////////////////
//////////////////////
//////////////////////
//////////////////////public class WordModifier {
//////////////////////    public static void main(String[] args) {
//////////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx"; // 配置文档的路径
//////////////////////
//////////////////////        File folder = new File("Z:\\Desktop\\测试\\in"); // 文件夹的路径
//////////////////////        File[] listOfFiles = folder.listFiles(); // 获取文件夹中的所有文件
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////        for (File file : listOfFiles) {
//////////////////////            if (file.isFile()) { // 检查文件是否是Word文档
//////////////////////                String sourceFile = file.getAbsolutePath(); // 获取文件的绝对路径
//////////////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName(); // 输出文档的路径
//////////////////////
//////////////////////                // 如果文件是.doc文件，将其转换为.docx
//////////////////////                if (file.getName().endsWith(".doc")) {
//////////////////////
//////////////////////                    String docxFile = sourceFile.replace(".doc", ".docx");
//////////////////////                                        // 使用Aspose Words的Document类进行转换
//////////////////////                                        com.aspose.words.Document doc = null;
//////////////////////                                        try {
//////////////////////                                                doc = new com.aspose.words.Document(sourceFile);
//////////////////////                                        } catch (Exception e) {
//////////////////////                                                throw new RuntimeException(e);
//////////////////////                                        }
//////////////////////                                        try {
//////////////////////                                                doc.save(docxFile, SaveFormat.DOCX);
//////////////////////                                        } catch (Exception e) {
//////////////////////                                                throw new RuntimeException(e);
//////////////////////                                        }
//////////////////////                                        sourceFile = docxFile; // 更新源文件路径为.docx文件
////////////////////////                    // 加载源 PDF 文件
////////////////////////                    Converter converter = new Converter(sourceFile);
////////////////////////
////////////////////////// 设置转换选项
////////////////////////                    WordProcessingConvertOptions convertOptions =
////////////////////////                            new WordProcessingConvertOptions();
////////////////////////
////////////////////////// 将 PDF 转换为 DOCX
////////////////////////                    converter.convert(docxFile, convertOptions);
//////////////////////                }
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////                // 确保源文件是.docx文件
//////////////////////                if (!sourceFile.endsWith(".docx")) {
//////////////////////                    continue;
//////////////////////                }
//////////////////////                Map<String, String> configMap = new HashMap<>(); // 创建一个映射来存储配置规则
//////////////////////
//////////////////////                // 读取配置文档
//////////////////////                try (FileInputStream configFis = new FileInputStream(configFile);
//////////////////////                     XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////////////
//////////////////////                    List<XWPFTable> configTables = configDoc.getTables(); // 获取配置文档中的所有表格
//////////////////////
//////////////////////                    // 遍历配置文档中的每个表格
//////////////////////                    for (XWPFTable table : configTables) {
//////////////////////                        // 遍历每个表格中的行
//////////////////////                        for (XWPFTableRow row : table.getRows()) {
//////////////////////                            // 假设每行的第一个单元格包含标签，第二个单元格包含更新的值
//////////////////////                            if (row.getTableCells().size() >= 2) {
//////////////////////                                XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////////////                                XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////////////                                String key = labelCell.getText().trim();
//////////////////////                                String value = valueCell.getText().trim();
//////////////////////                                configMap.put(key, value); // 将键值对添加到映射中
//////////////////////                            }
//////////////////////                        }
//////////////////////                    }
//////////////////////                } catch (IOException e) {
//////////////////////                    e.printStackTrace();
//////////////////////                }
//////////////////////
//////////////////////                // 打印配置映射，调试用
//////////////////////                System.out.println("配置映射：");
//////////////////////                for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////////////                                        System.out.println(entry.getKey() + " => " + entry.getValue());
//////////////////////                }
//////////////////////
//////////////////////                // 读取和修改目标文档
//////////////////////                try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////////////                     FileOutputStream fos = new FileOutputStream(outputFile);
//////////////////////                     XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////////////
//////////////////////                    List<XWPFTable> tables = document.getTables(); // 获取源文档中的所有表格
//////////////////////
//////////////////////                    // 遍历源文档中的每个表格
//////////////////////                    for (XWPFTable table : tables) {
//////////////////////                        // 遍历每个表格中的行
//////////////////////                        for (XWPFTableRow row : table.getRows()) {
//////////////////////                            // 遍历行中的每个单元格
//////////////////////                            for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////////////                                XWPFTableCell cell = row.getTableCells().get(i);
//////////////////////                                String text = cell.getText().replaceAll("\\s+", ""); // 获取单元格中的文本
//////////////////////
//////////////////////                                // 打印调试信息
////////////////////////                                                                System.out.println("表格单元格文本：" + text);
//////////////////////
//////////////////////                                // 如果单元格的文本存在于配置规则的映射中
//////////////////////                                if (configMap.containsKey(text)) {
//////////////////////                                    String newValue = configMap.get(text);
//////////////////////                                    // 检查并更新下一个单元格（如果存在和有新值）
//////////////////////                                    if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////////////                                        XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////////////
//////////////////////                                        // 清空单元格
//////////////////////                                        nextCell.removeParagraph(0);
//////////////////////
//////////////////////                                        // 添加新的内容
//////////////////////                                        XWPFParagraph p = nextCell.addParagraph();
//////////////////////                                        p.setAlignment(ParagraphAlignment.CENTER); // 设置段落为居中对齐
//////////////////////                                        XWPFRun r = p.createRun();
//////////////////////                                        r.setText(newValue);
//////////////////////
//////////////////////                                        nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER); // 设置单元格内容为水平居中
//////////////////////                                    }
//////////////////////                                }
//////////////////////                            }
//////////////////////                        }
//////////////////////                    }
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////                    // 遍历文档中的所有段落
//////////////////////                    for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////////////                        List<XWPFRun> runs = paragraph.getRuns();
//////////////////////                        for (int i = 0; i < runs.size(); i++) {
//////////////////////                            XWPFRun run = runs.get(i);
//////////////////////                            String text = run.getText(0);
//////////////////////                            if (text != null) {
//////////////////////                                for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////////////                                    String key = entry.getKey();
//////////////////////                                    String value = entry.getValue();
//////////////////////                                    if(key.equals(text.trim())){
//////////////////////                                        System.out.println(key+" **   "+text.trim()+"      是否相等  "+key.equals(text.trim()));
//////////////////////                                    }
//////////////////////                                    if (text.trim().equals(key)) { // 检查文本是否包含键和一个空格和一个冒号
//////////////////////                                        // 修改后面的文本
//////////////////////                                        int j = i + 1;
//////////////////////                                        while (j < runs.size()) {
//////////////////////                                            XWPFRun nextRun = runs.get(j);
//////////////////////                                            String nextText = nextRun.getText(0);
//////////////////////                                            System.out.println("########");
//////////////////////                                            System.out.println(nextText);
//////////////////////                                            if (nextText != null && !nextText.contains(":")) {
//////////////////////                                                // 创建新的 run 并设置文本
//////////////////////                                                if (j+1 <= runs.size()) {
//////////////////////                                                    XWPFRun newRun = paragraph.insertNewRun(j+1);
//////////////////////                                                    newRun.setText(value+"   ");
//////////////////////
//////////////////////                                                    // 给新的 run 添加下划线
//////////////////////                                                    newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////////////
//////////////////////                                                    // 设置字体和字号
//////////////////////                                                    newRun.setFontFamily("仿宋_GB2312");
//////////////////////                                                    newRun.setFontSize(14); // 四号字体对应的字号大约为14pt
//////////////////////                                                    // 删除旧的 run
//////////////////////                                                    paragraph.removeRun(j);
//////////////////////                                                }
//////////////////////
//////////////////////                                                break;
//////////////////////                                            }
//////////////////////
//////////////////////
//////////////////////
//////////////////////                                            j++;
//////////////////////                                        }
//////////////////////                                        i = j; // 跳过已经修改的运行
//////////////////////                                        break; // 找到一个匹配的键后，就退出循环
//////////////////////                                    }
//////////////////////                                }
//////////////////////                            }
//////////////////////                        }
//////////////////////                    }
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////
//////////////////////                    document.write(fos); // 将修改后的文档写入到输出文件中
//////////////////////                } catch (IOException e) {
//////////////////////                    e.printStackTrace();
//////////////////////                }
//////////////////////
//////////////////////
//////////////////////
//////////////////////            }
//////////////////////        }
//////////////////////    }
//////////////////////
//////////////////////
//////////////////////
//////////////////////}
//////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////
////////////////////package org.example;
////////////////////
////////////////////import com.aspose.words.SaveFormat;
////////////////////import org.apache.poi.xwpf.usermodel.*;
////////////////////import javax.swing.*;
////////////////////import javax.swing.table.DefaultTableModel;
////////////////////import org.jfree.chart.ChartFactory;
////////////////////import org.jfree.chart.ChartPanel;
////////////////////import org.jfree.chart.JFreeChart;
////////////////////import org.jfree.chart.plot.PlotOrientation;
////////////////////import org.jfree.data.category.DefaultCategoryDataset;
////////////////////
////////////////////import java.awt.*;
////////////////////import java.io.File;
////////////////////import java.io.FileInputStream;
////////////////////import java.io.FileOutputStream;
////////////////////import java.io.IOException;
////////////////////import java.util.HashMap;
////////////////////import java.util.List;
////////////////////import java.util.Map;
////////////////////
////////////////////public class WordModifier {
////////////////////    private static Map<String, String> configMap = new HashMap<>();
////////////////////    private static DefaultTableModel tableModel = new DefaultTableModel(new Object[]{"文件", "错误类型", "描述"}, 0);
////////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////////////    private static DefaultCategoryDataset dataset = new DefaultCategoryDataset();
////////////////////
////////////////////    public static void main(String[] args) {
////////////////////        // 创建并显示主窗口
////////////////////        JFrame frame = createMainFrame();
////////////////////        frame.setVisible(true);
////////////////////
////////////////////        // 加载配置文件
////////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////////////        loadConfigFile(configFile);
////////////////////
////////////////////        // 处理文档文件夹中的文件
////////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////////////        File[] listOfFiles = folder.listFiles();
////////////////////
////////////////////        if (listOfFiles != null) {
////////////////////            int totalFiles = listOfFiles.length;
////////////////////            int processedFiles = 0;
////////////////////            int successCount = 0;
////////////////////            int failureCount = 0;
////////////////////
////////////////////            for (File file : listOfFiles) {
////////////////////                if (file.isFile()) {
////////////////////                    String sourceFile = file.getAbsolutePath();
////////////////////                    String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////////////
////////////////////                    // 如果文件是.doc文件，将其转换为.docx
////////////////////                    if (sourceFile.endsWith(".doc")) {
////////////////////                        sourceFile = convertDocToDocx(sourceFile);
////////////////////                    }
////////////////////
////////////////////                    // 确保源文件是.docx文件
////////////////////                    if (!sourceFile.endsWith(".docx")) {
////////////////////                        continue;
////////////////////                    }
////////////////////
////////////////////                    try {
////////////////////                        modifyDocument(sourceFile, outputFile);
////////////////////                        successCount++;
////////////////////                    } catch (Exception e) {
////////////////////                        tableModel.addRow(new Object[]{file.getName(), "处理错误", e.getMessage()});
////////////////////                        failureCount++;
////////////////////                    }
////////////////////
////////////////////                    processedFiles++;
////////////////////                    updateProgress(processedFiles, totalFiles);
////////////////////                }
////////////////////            }
////////////////////
////////////////////            dataset.addValue(successCount, "数量", "成功");
////////////////////            dataset.addValue(failureCount, "数量", "失败");
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static void loadConfigFile(String configFile) {
////////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////////////
////////////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////////////
////////////////////            for (XWPFTable table : configTables) {
////////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////////                    if (row.getTableCells().size() >= 2) {
////////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////////////                        String key = labelCell.getText().trim();
////////////////////                        String value = valueCell.getText().trim();
////////////////////                        configMap.put(key, value);
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////        } catch (IOException e) {
////////////////////            e.printStackTrace();
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static String convertDocToDocx(String sourceFile) {
////////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////////////        try {
////////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////////////        } catch (Exception e) {
////////////////////            throw new RuntimeException(e);
////////////////////        }
////////////////////        return docxFile;
////////////////////    }
////////////////////
////////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////////////
////////////////////            List<XWPFTable> tables = document.getTables();
////////////////////
////////////////////            for (XWPFTable table : tables) {
////////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////////////
////////////////////                        if (configMap.containsKey(text)) {
////////////////////                            String newValue = configMap.get(text);
////////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////////////                                nextCell.removeParagraph(0);
////////////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////////////                                XWPFRun r = p.createRun();
////////////////////                                r.setText(newValue);
////////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////////////                            }
////////////////////                        }
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////
////////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////////////                for (int i = 0; i < runs.size(); i++) {
////////////////////                    XWPFRun run = runs.get(i);
////////////////////                    String text = run.getText(0);
////////////////////                    if (text != null) {
////////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////////                            String key = entry.getKey();
////////////////////                            String value = entry.getValue();
////////////////////                            if (text.trim().equals(key)) {
////////////////////                                int j = i + 1;
////////////////////                                while (j < runs.size()) {
////////////////////                                    XWPFRun nextRun = runs.get(j);
////////////////////                                    String nextText = nextRun.getText(0);
////////////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////////////                                        newRun.setText(value);
////////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////////////                                        newRun.setFontSize(14);
////////////////////                                        paragraph.removeRun(j);
////////////////////                                        break;
////////////////////                                    }
////////////////////                                    j++;
////////////////////                                }
////////////////////                                i = j;
////////////////////                                break;
////////////////////                            }
////////////////////                        }
////////////////////                    }
////////////////////                }
////////////////////            }
////////////////////
////////////////////            document.write(fos);
////////////////////        }
////////////////////    }
////////////////////
////////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////////////        progressBar.setValue(progress);
////////////////////    }
////////////////////
////////////////////    private static JFrame createMainFrame() {
////////////////////        JFrame frame = new JFrame("文档处理工具");
////////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////////////        frame.setSize(800, 600);
////////////////////
////////////////////        JPanel panel = new JPanel(new BorderLayout());
////////////////////        frame.add(panel);
////////////////////
////////////////////        JTable table = new JTable(tableModel);
////////////////////        panel.add(new JScrollPane(table), BorderLayout.NORTH);
////////////////////
////////////////////        panel.add(progressBar, BorderLayout.CENTER);
////////////////////
////////////////////        JFreeChart barChart = ChartFactory.createBarChart(
////////////////////                "文档处理统计分析",
////////////////////                "类别",
////////////////////                "数量",
////////////////////                dataset,
////////////////////                PlotOrientation.VERTICAL,
////////////////////                true, true, false);
////////////////////        ChartPanel chartPanel = new ChartPanel(barChart);
////////////////////        chartPanel.setPreferredSize(new Dimension(800, 400));
////////////////////        panel.add(chartPanel, BorderLayout.SOUTH);
////////////////////
////////////////////        return frame;
////////////////////    }
////////////////////}
////////////////////
////////////////////
////////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////
//////////////////package org.example;
//////////////////
//////////////////import com.aspose.words.SaveFormat;
//////////////////import org.apache.poi.xwpf.usermodel.*;
//////////////////import javax.swing.*;
//////////////////import javax.swing.table.DefaultTableModel;
//////////////////import org.jfree.chart.ChartFactory;
//////////////////import org.jfree.chart.ChartPanel;
//////////////////import org.jfree.chart.JFreeChart;
//////////////////import org.jfree.chart.plot.PlotOrientation;
//////////////////import org.jfree.data.category.DefaultCategoryDataset;
//////////////////
//////////////////import java.awt.*;
//////////////////import java.awt.event.ActionEvent;
//////////////////import java.awt.event.ActionListener;
//////////////////import java.io.File;
//////////////////import java.io.FileInputStream;
//////////////////import java.io.FileOutputStream;
//////////////////import java.io.IOException;
//////////////////import java.util.HashMap;
//////////////////import java.util.List;
//////////////////import java.util.Map;
//////////////////
//////////////////public class WordModifier {
//////////////////    private static Map<String, String> configMap = new HashMap<>();
//////////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////////////    private static DefaultCategoryDataset dataset = new DefaultCategoryDataset();
//////////////////    private static JFrame frame;
//////////////////    private static JTextField keyField = new JTextField();
//////////////////    private static JTextField valueField = new JTextField();
//////////////////    private static JTable configTable;
//////////////////    private static JTable fileTable;
//////////////////
//////////////////    public static void main(String[] args) {
//////////////////        // 创建并显示主窗口
//////////////////        frame = createMainFrame();
//////////////////        frame.setVisible(true);
//////////////////
//////////////////        // 加载配置文件
//////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////////////        loadConfigFile(configFile);
//////////////////
//////////////////        // 处理文档文件夹中的文件
//////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////////////        File[] listOfFiles = folder.listFiles();
//////////////////
//////////////////        if (listOfFiles != null) {
//////////////////            int totalFiles = listOfFiles.length;
//////////////////            int processedFiles = 0;
//////////////////            int successCount = 0;
//////////////////            int failureCount = 0;
//////////////////
//////////////////            for (File file : listOfFiles) {
//////////////////                if (file.isFile()) {
//////////////////                    String sourceFile = file.getAbsolutePath();
//////////////////                    String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////////////
//////////////////                    // 如果文件是.doc文件，将其转换为.docx
//////////////////                    if (sourceFile.endsWith(".doc")) {
//////////////////                        sourceFile = convertDocToDocx(sourceFile);
//////////////////                    }
//////////////////
//////////////////                    // 确保源文件是.docx文件
//////////////////                    if (!sourceFile.endsWith(".docx")) {
//////////////////                        continue;
//////////////////                    }
//////////////////
//////////////////                    try {
//////////////////                        modifyDocument(sourceFile, outputFile);
//////////////////                        successCount++;
//////////////////                    } catch (Exception e) {
//////////////////                        fileTableModel.addRow(new Object[]{file.getName(), "处理错误", e.getMessage()});
//////////////////                        failureCount++;
//////////////////                    }
//////////////////
//////////////////                    processedFiles++;
//////////////////                    updateProgress(processedFiles, totalFiles);
//////////////////                }
//////////////////            }
//////////////////
//////////////////            dataset.addValue(successCount, "数量", "成功");
//////////////////            dataset.addValue(failureCount, "数量", "失败");
//////////////////        }
//////////////////    }
//////////////////
//////////////////    private static void loadConfigFile(String configFile) {
//////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////////
//////////////////            List<XWPFTable> configTables = configDoc.getTables();
//////////////////
//////////////////            for (XWPFTable table : configTables) {
//////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////                    if (row.getTableCells().size() >= 2) {
//////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////////                        String key = labelCell.getText().trim();
//////////////////                        String value = valueCell.getText().trim();
//////////////////                        configMap.put(key, value);
//////////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////        } catch (IOException e) {
//////////////////            e.printStackTrace();
//////////////////        }
//////////////////    }
//////////////////
//////////////////    private static String convertDocToDocx(String sourceFile) {
//////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////////////        try {
//////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////////////            doc.save(docxFile, SaveFormat.DOCX);
//////////////////        } catch (Exception e) {
//////////////////            throw new RuntimeException(e);
//////////////////        }
//////////////////        return docxFile;
//////////////////    }
//////////////////
//////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////////
//////////////////            StringBuilder originalContent = new StringBuilder();
//////////////////            StringBuilder modifiedContent = new StringBuilder();
//////////////////
//////////////////            List<XWPFTable> tables = document.getTables();
//////////////////
//////////////////            for (XWPFTable table : tables) {
//////////////////                for (XWPFTableRow row : table.getRows()) {
//////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////////////                        originalContent.append(text).append(" ");
//////////////////
//////////////////                        if (configMap.containsKey(text)) {
//////////////////                            String newValue = configMap.get(text);
//////////////////                            modifiedContent.append(newValue).append(" ");
//////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////////                                nextCell.removeParagraph(0);
//////////////////                                XWPFParagraph p = nextCell.addParagraph();
//////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////////////                                XWPFRun r = p.createRun();
//////////////////                                r.setText(newValue);
//////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////////////                            }
//////////////////                        }
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////
//////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////////                List<XWPFRun> runs = paragraph.getRuns();
//////////////////                for (int i = 0; i < runs.size(); i++) {
//////////////////                    XWPFRun run = runs.get(i);
//////////////////                    String text = run.getText(0);
//////////////////                    if (text != null) {
//////////////////                        originalContent.append(text).append(" ");
//////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////////                            String key = entry.getKey();
//////////////////                            String value = entry.getValue();
//////////////////                            if (text.trim().equals(key)) {
//////////////////                                int j = i + 1;
//////////////////                                while (j < runs.size()) {
//////////////////                                    XWPFRun nextRun = runs.get(j);
//////////////////                                    String nextText = nextRun.getText(0);
//////////////////                                    if (nextText != null && !nextText.contains(":")) {
//////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////////////                                        newRun.setText(value);
//////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////////////                                        newRun.setFontSize(14);
//////////////////                                        paragraph.removeRun(j);
//////////////////                                        break;
//////////////////                                    }
//////////////////                                    j++;
//////////////////                                }
//////////////////                                i = j;
//////////////////                                break;
//////////////////                            }
//////////////////                        }
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////
//////////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////////////            document.write(fos);
//////////////////        }
//////////////////    }
//////////////////
//////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////////////        progressBar.setValue(progress);
//////////////////    }
//////////////////
//////////////////    private static JFrame createMainFrame() {
//////////////////        JFrame frame = new JFrame("文档处理工具");
//////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////////////        frame.setSize(1000, 800);
//////////////////
//////////////////        JPanel panel = new JPanel(new BorderLayout());
//////////////////        frame.add(panel);
//////////////////
//////////////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////////////        configTable = new JTable(configTableModel);
//////////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////////////
//////////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////////////        configInputPanel.add(new JLabel("Key:"));
//////////////////        configInputPanel.add(keyField);
//////////////////        configInputPanel.add(new JLabel("Value:"));
//////////////////        configInputPanel.add(valueField);
//////////////////
//////////////////        JButton addButton = new JButton("添加/更新");
//////////////////        addButton.addActionListener(new ActionListener() {
//////////////////            @Override
//////////////////            public void actionPerformed(ActionEvent e) {
//////////////////                String key = keyField.getText().trim();
//////////////////                String value = valueField.getText().trim();
//////////////////                if (!key.isEmpty() && !value.isEmpty()) {
//////////////////                    configMap.put(key, value);
//////////////////                    boolean keyExists = false;
//////////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////////////                            configTableModel.setValueAt(value, i, 1);
//////////////////                            keyExists = true;
//////////////////                            break;
//////////////////                        }
//////////////////                    }
//////////////////                    if (!keyExists) {
//////////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////////                    }
//////////////////                }
//////////////////            }
//////////////////        });
//////////////////        configInputPanel.add(addButton);
//////////////////
//////////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////////////
//////////////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果"));
//////////////////        fileTable = new JTable(fileTableModel);
//////////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////////////
//////////////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////////////
//////////////////        JPanel chartPanel = new JPanel(new BorderLayout());
//////////////////        chartPanel.setBorder(BorderFactory.createTitledBorder("文档处理统计分析"));
//////////////////        JFreeChart barChart = ChartFactory.createBarChart(
//////////////////                "文档处理统计分析",
//////////////////                "类别",
//////////////////                "数量",
//////////////////                dataset,
//////////////////                PlotOrientation.VERTICAL,
//////////////////                true, true, false);
//////////////////        ChartPanel chartPanelInner = new ChartPanel(barChart);
//////////////////        chartPanel.add(chartPanelInner, BorderLayout.CENTER);
//////////////////
//////////////////        panel.add(configPanel, BorderLayout.NORTH);
//////////////////        panel.add(filePanel, BorderLayout.CENTER);
//////////////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////////////        panel.add(chartPanel, BorderLayout.EAST);
//////////////////
//////////////////        return frame;
//////////////////    }
//////////////////}
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////
////////////////package org.example;
////////////////
////////////////import com.aspose.words.SaveFormat;
////////////////import org.apache.poi.xwpf.usermodel.*;
////////////////import javax.swing.*;
////////////////import javax.swing.table.DefaultTableModel;
////////////////import java.awt.*;
////////////////import java.awt.event.ActionEvent;
////////////////import java.awt.event.ActionListener;
////////////////import java.io.File;
////////////////import java.io.FileInputStream;
////////////////import java.io.FileOutputStream;
////////////////import java.io.IOException;
////////////////import java.util.HashMap;
////////////////import java.util.List;
////////////////import java.util.Map;
////////////////
////////////////public class WordModifier {
////////////////    private static Map<String, String> configMap = new HashMap<>();
////////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////////    private static JFrame frame;
////////////////    private static JTextField keyField = new JTextField();
////////////////    private static JTextField valueField = new JTextField();
////////////////    private static JTable configTable;
////////////////    private static JTable fileTable;
////////////////    private static JTextArea statsTextArea = new JTextArea();
////////////////
////////////////    public static void main(String[] args) {
////////////////        // 创建并显示主窗口
////////////////        frame = createMainFrame();
////////////////        frame.setVisible(true);
////////////////
////////////////        // 加载配置文件
////////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////////        loadConfigFile(configFile);
////////////////    }
////////////////
////////////////    private static void loadConfigFile(String configFile) {
////////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////////
////////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////////
////////////////            for (XWPFTable table : configTables) {
////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////                    if (row.getTableCells().size() >= 2) {
////////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////////                        String key = labelCell.getText().trim();
////////////////                        String value = valueCell.getText().trim();
////////////////                        configMap.put(key, value);
////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////                    }
////////////////                }
////////////////            }
////////////////        } catch (IOException e) {
////////////////            e.printStackTrace();
////////////////        }
////////////////    }
////////////////
////////////////    private static String convertDocToDocx(String sourceFile) {
////////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////////        try {
////////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////////        } catch (Exception e) {
////////////////            throw new RuntimeException(e);
////////////////        }
////////////////        return docxFile;
////////////////    }
////////////////
////////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////////
////////////////            StringBuilder originalContent = new StringBuilder();
////////////////            StringBuilder modifiedContent = new StringBuilder();
////////////////
////////////////            List<XWPFTable> tables = document.getTables();
////////////////
////////////////            for (XWPFTable table : tables) {
////////////////                for (XWPFTableRow row : table.getRows()) {
////////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////////                        originalContent.append(text).append(" ");
////////////////
////////////////                        if (configMap.containsKey(text)) {
////////////////                            String newValue = configMap.get(text);
////////////////                            modifiedContent.append(newValue).append(" ");
////////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////////                                nextCell.removeParagraph(0);
////////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////////                                XWPFRun r = p.createRun();
////////////////                                r.setText(newValue);
////////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////////                            }
////////////////                        }
////////////////                    }
////////////////                }
////////////////            }
////////////////
////////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////////                for (int i = 0; i < runs.size(); i++) {
////////////////                    XWPFRun run = runs.get(i);
////////////////                    String text = run.getText(0);
////////////////                    if (text != null) {
////////////////                        originalContent.append(text).append(" ");
////////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////////                            String key = entry.getKey();
////////////////                            String value = entry.getValue();
////////////////                            if (text.trim().equals(key)) {
////////////////                                int j = i + 1;
////////////////                                while (j < runs.size()) {
////////////////                                    XWPFRun nextRun = runs.get(j);
////////////////                                    String nextText = nextRun.getText(0);
////////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////////                                        newRun.setText(value);
////////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////////                                        newRun.setFontSize(14);
////////////////                                        paragraph.removeRun(j);
////////////////                                        break;
////////////////                                    }
////////////////                                    j++;
////////////////                                }
////////////////                                i = j;
////////////////                                break;
////////////////                            }
////////////////                        }
////////////////                    }
////////////////                }
////////////////            }
////////////////
////////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////////////            document.write(fos);
////////////////        }
////////////////    }
////////////////
////////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////////        progressBar.setValue(progress);
////////////////    }
////////////////
////////////////    private static JFrame createMainFrame() {
////////////////        JFrame frame = new JFrame("文档处理工具");
////////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////////        frame.setSize(1200, 800);
////////////////
////////////////        JPanel panel = new JPanel(new BorderLayout());
////////////////        frame.add(panel);
////////////////
////////////////        JPanel configPanel = new JPanel(new BorderLayout());
////////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////////////        configTable = new JTable(configTableModel);
////////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////////////
////////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////////////        configInputPanel.add(new JLabel("Key:"));
////////////////        configInputPanel.add(keyField);
////////////////        configInputPanel.add(new JLabel("Value:"));
////////////////        configInputPanel.add(valueField);
////////////////
////////////////        JButton addButton = new JButton("添加/更新");
////////////////        addButton.addActionListener(new ActionListener() {
////////////////            @Override
////////////////            public void actionPerformed(ActionEvent e) {
////////////////                String key = keyField.getText().trim();
////////////////                String value = valueField.getText().trim();
////////////////                if (!key.isEmpty() && !value.isEmpty()) {
////////////////                    configMap.put(key, value);
////////////////                    boolean keyExists = false;
////////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////////////                            configTableModel.setValueAt(value, i, 1);
////////////////                            keyExists = true;
////////////////                            break;
////////////////                        }
////////////////                    }
////////////////                    if (!keyExists) {
////////////////                        configTableModel.addRow(new Object[]{key, value});
////////////////                    }
////////////////                }
////////////////            }
////////////////        });
////////////////        configInputPanel.add(addButton);
////////////////
////////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////////////
////////////////        JPanel filePanel = new JPanel(new BorderLayout());
////////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////////////        fileTable = new JTable(fileTableModel);
////////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
////////////////            if (!event.getValueIsAdjusting()) {
////////////////                int selectedRow = fileTable.getSelectedRow();
////////////////                if (selectedRow >= 0) {
////////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
////////////////                }
////////////////            }
////////////////        });
////////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////////////
////////////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////////////
////////////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("文档处理统计分析"));
////////////////        statsTextArea.setEditable(false);
////////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////////////
////////////////        JButton startButton = new JButton("开始执行");
////////////////        startButton.addActionListener(new ActionListener() {
////////////////            @Override
////////////////            public void actionPerformed(ActionEvent e) {
////////////////                processFiles();
////////////////            }
////////////////        });
////////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////////////
////////////////        panel.add(configPanel, BorderLayout.NORTH);
////////////////        panel.add(filePanel, BorderLayout.CENTER);
////////////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////////////        panel.add(statsPanel, BorderLayout.EAST);
////////////////
////////////////        return frame;
////////////////    }
////////////////
////////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////////////        previewFrame.setSize(600, 400);
////////////////        JTextArea previewTextArea = new JTextArea();
////////////////        previewTextArea.setEditable(false);
////////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////////////        previewFrame.add(new JScrollPane(previewTextArea));
////////////////        previewFrame.setVisible(true);
////////////////    }
////////////////
////////////////    private static void processFiles() {
////////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////////        File[] listOfFiles = folder.listFiles();
////////////////        if (listOfFiles == null) {
////////////////            return;
////////////////        }
////////////////        int totalFiles = listOfFiles.length;
////////////////        int processedFiles = 0;
////////////////        for (File file : listOfFiles) {
////////////////            if (file.isFile()) {
////////////////                String sourceFile = file.getAbsolutePath();
////////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////////                if (sourceFile.endsWith(".doc")) {
////////////////                    sourceFile = convertDocToDocx(sourceFile);
////////////////                }
////////////////                if (!sourceFile.endsWith(".docx")) {
////////////////                    continue;
////////////////                }
////////////////                try {
////////////////                    modifyDocument(sourceFile, outputFile);
////////////////                    processedFiles++;
////////////////                    updateProgress(processedFiles, totalFiles);
////////////////                } catch (IOException e) {
////////////////                    e.printStackTrace();
////////////////                }
////////////////            }
////////////////        }
////////////////        displayStats(totalFiles, processedFiles);
////////////////    }
////////////////
////////////////    private static void displayStats(int totalFiles, int processedFiles) {
////////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n";
////////////////        statsTextArea.setText(stats);
////////////////    }
////////////////}
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////
//////////////package org.example;
//////////////
//////////////import com.aspose.words.SaveFormat;
//////////////import org.apache.poi.xwpf.usermodel.*;
//////////////
//////////////import javax.swing.*;
//////////////import javax.swing.table.DefaultTableModel;
//////////////import java.awt.*;
//////////////import java.awt.event.ActionEvent;
//////////////import java.awt.event.ActionListener;
//////////////import java.io.File;
//////////////import java.io.FileInputStream;
//////////////import java.io.FileOutputStream;
//////////////import java.io.IOException;
//////////////import java.util.HashMap;
//////////////import java.util.List;
//////////////import java.util.Map;
//////////////
//////////////public class WordModifier {
//////////////    private static Map<String, String> configMap = new HashMap<>();
//////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////////    private static JFrame frame;
//////////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////////////    private static JTextField valueField = new JTextField();
//////////////    private static JTable configTable;
//////////////    private static JTable fileTable;
//////////////    private static JTextArea statsTextArea = new JTextArea();
//////////////
//////////////    public static void main(String[] args) {
//////////////        // 创建并显示主窗口
//////////////        frame = createMainFrame();
//////////////        frame.setVisible(true);
//////////////
//////////////        // 加载配置文件
//////////////        String configFile = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////////        loadConfigFile(configFile);
//////////////    }
//////////////
//////////////    private static void loadConfigFile(String configFile) {
//////////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////////
//////////////            List<XWPFTable> configTables = configDoc.getTables();
//////////////
//////////////            for (XWPFTable table : configTables) {
//////////////                for (XWPFTableRow row : table.getRows()) {
//////////////                    if (row.getTableCells().size() >= 2) {
//////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////////                        String key = labelCell.getText().trim();
//////////////                        String value = valueCell.getText().trim();
//////////////                        configMap.put(key, value);
//////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////                        keyComboBox.addItem(key);
//////////////                    }
//////////////                }
//////////////            }
//////////////        } catch (IOException e) {
//////////////            e.printStackTrace();
//////////////        }
//////////////    }
//////////////
//////////////    private static void saveConfigFile(String configFile) {
//////////////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////////////             XWPFDocument configDoc = new XWPFDocument()) {
//////////////
//////////////            XWPFTable table = configDoc.createTable();
//////////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////                XWPFTableRow row = table.createRow();
//////////////                row.getCell(0).setText(entry.getKey());
//////////////                row.getCell(1).setText(entry.getValue());
//////////////            }
//////////////            configDoc.write(fos);
//////////////        } catch (IOException e) {
//////////////            e.printStackTrace();
//////////////        }
//////////////    }
//////////////
//////////////    private static String convertDocToDocx(String sourceFile) {
//////////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////////        try {
//////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////////            doc.save(docxFile, SaveFormat.DOCX);
//////////////        } catch (Exception e) {
//////////////            throw new RuntimeException(e);
//////////////        }
//////////////        return docxFile;
//////////////    }
//////////////
//////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////////
//////////////            StringBuilder originalContent = new StringBuilder();
//////////////            StringBuilder modifiedContent = new StringBuilder();
//////////////
//////////////            List<XWPFTable> tables = document.getTables();
//////////////
//////////////            for (XWPFTable table : tables) {
//////////////                for (XWPFTableRow row : table.getRows()) {
//////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////////                        originalContent.append(text).append(" ");
//////////////
//////////////                        if (configMap.containsKey(text)) {
//////////////                            String newValue = configMap.get(text);
//////////////                            modifiedContent.append(newValue).append(" ");
//////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////////                                nextCell.removeParagraph(0);
//////////////                                XWPFParagraph p = nextCell.addParagraph();
//////////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////////                                XWPFRun r = p.createRun();
//////////////                                r.setText(newValue);
//////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////////                            }
//////////////                        }
//////////////                    }
//////////////                }
//////////////            }
//////////////
//////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////////                List<XWPFRun> runs = paragraph.getRuns();
//////////////                for (int i = 0; i < runs.size(); i++) {
//////////////                    XWPFRun run = runs.get(i);
//////////////                    String text = run.getText(0);
//////////////                    if (text != null) {
//////////////                        originalContent.append(text).append(" ");
//////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////////                            String key = entry.getKey();
//////////////                            String value = entry.getValue();
//////////////                            if (text.trim().equals(key)) {
//////////////                                int j = i + 1;
//////////////                                while (j < runs.size()) {
//////////////                                    XWPFRun nextRun = runs.get(j);
//////////////                                    String nextText = nextRun.getText(0);
//////////////                                    if (nextText != null && !nextText.contains(":")) {
//////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////////                                        newRun.setText(value);
//////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////////                                        newRun.setFontSize(14);
//////////////                                        paragraph.removeRun(j);
//////////////                                        break;
//////////////                                    }
//////////////                                    j++;
//////////////                                }
//////////////                                i = j;
//////////////                                break;
//////////////                            }
//////////////                        }
//////////////                    }
//////////////                }
//////////////            }
//////////////
//////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////////            document.write(fos);
//////////////        }
//////////////    }
//////////////
//////////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////////        progressBar.setValue(progress);
//////////////    }
//////////////
//////////////    private static JFrame createMainFrame() {
//////////////        JFrame frame = new JFrame("文档处理工具");
//////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////////        frame.setSize(1200, 800);
//////////////
//////////////        JPanel panel = new JPanel(new BorderLayout());
//////////////        frame.add(panel);
//////////////
//////////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////////        configTable = new JTable(configTableModel);
//////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////////
//////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////////        configInputPanel.add(new JLabel("Key:"));
//////////////        configInputPanel.add(keyComboBox);
//////////////        configInputPanel.add(new JLabel("Value:"));
//////////////        configInputPanel.add(valueField);
//////////////
//////////////        JButton addButton = new JButton("添加/更新");
//////////////        addButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                String key = (String) keyComboBox.getSelectedItem();
//////////////                String value = valueField.getText().trim();
//////////////                if (!key.isEmpty() && !value.isEmpty()) {
//////////////                    configMap.put(key, value);
//////////////                    boolean keyExists = false;
//////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////////                            configTableModel.setValueAt(value, i, 1);
//////////////                            keyExists = true;
//////////////                            break;
//////////////                        }
//////////////                    }
//////////////                    if (!keyExists) {
//////////////                        configTableModel.addRow(new Object[]{key, value});
//////////////                        keyComboBox.addItem(key);
//////////////                    }
//////////////                    saveConfigFile("Z:\\Desktop\\测试\\模板\\模板.docx");
//////////////                }
//////////////            }
//////////////        });
//////////////        configInputPanel.add(addButton);
//////////////
//////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////////
//////////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////////////        fileTable = new JTable(fileTableModel);
//////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
//////////////            if (!event.getValueIsAdjusting()) {
//////////////                int selectedRow = fileTable.getSelectedRow();
//////////////                if (selectedRow != -1) {
//////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
//////////////                }
//////////////            }
//////////////        });
//////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////////
//////////////        JButton refreshButton = new JButton("刷新预览");
//////////////        refreshButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                fileTableModel.setRowCount(0);
//////////////                processFiles();
//////////////            }
//////////////        });
//////////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////////////
//////////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////////
//////////////        JButton startButton = new JButton("开始执行");
//////////////        startButton.addActionListener(new ActionListener() {
//////////////            @Override
//////////////            public void actionPerformed(ActionEvent e) {
//////////////                processFiles();
//////////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////////////            }
//////////////        });
//////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////////////
//////////////        JPanel statsPanel = new JPanel(new BorderLayout());
//////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////////////        statsTextArea.setEditable(false);
//////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////////////
//////////////        panel.add(configPanel, BorderLayout.NORTH);
//////////////        panel.add(filePanel, BorderLayout.CENTER);
//////////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////////        panel.add(statsPanel, BorderLayout.EAST);
//////////////
//////////////        return frame;
//////////////    }
//////////////
//////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////////////        previewFrame.setSize(600, 400);
//////////////        JTextArea previewTextArea = new JTextArea();
//////////////        previewTextArea.setEditable(false);
//////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////////////        previewFrame.add(new JScrollPane(previewTextArea));
//////////////        previewFrame.setVisible(true);
//////////////    }
//////////////
//////////////    private static void processFiles() {
//////////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////////        File[] listOfFiles = folder.listFiles();
//////////////        if (listOfFiles == null) {
//////////////            return;
//////////////        }
//////////////        int totalFiles = listOfFiles.length;
//////////////        int processedFiles = 0;
//////////////        long startTime = System.currentTimeMillis();
//////////////        for (File file : listOfFiles) {
//////////////            if (file.isFile()) {
//////////////                String sourceFile = file.getAbsolutePath();
//////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////////                if (sourceFile.endsWith(".doc")) {
//////////////                    sourceFile = convertDocToDocx(sourceFile);
//////////////                }
//////////////                if (!sourceFile.endsWith(".docx")) {
//////////////                    continue;
//////////////                }
//////////////                try {
//////////////                    modifyDocument(sourceFile, outputFile);
//////////////                    processedFiles++;
//////////////                    updateProgress(processedFiles, totalFiles);
//////////////                } catch (IOException e) {
//////////////                    e.printStackTrace();
//////////////                }
//////////////            }
//////////////        }
//////////////        long endTime = System.currentTimeMillis();
//////////////        long duration = endTime - startTime;
//////////////        displayStats(totalFiles, processedFiles, duration);
//////////////    }
//////////////
//////////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////////////        statsTextArea.setText(stats);
//////////////    }
//////////////}
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////
////////////package org.example;
////////////
////////////import com.aspose.words.SaveFormat;
////////////import org.apache.poi.xwpf.usermodel.*;
////////////
////////////import javax.swing.*;
////////////import javax.swing.table.DefaultTableModel;
////////////import java.awt.*;
////////////import java.awt.event.ActionEvent;
////////////import java.awt.event.ActionListener;
////////////import java.io.File;
////////////import java.io.FileInputStream;
////////////import java.io.FileOutputStream;
////////////import java.io.IOException;
////////////import java.util.HashMap;
////////////import java.util.List;
////////////import java.util.Map;
////////////
////////////public class WordModifier {
////////////    private static Map<String, String> configMap = new HashMap<>();
////////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////////    private static JFrame frame;
////////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////////////    private static JTextField valueField = new JTextField();
////////////    private static JTable configTable;
////////////    private static JTable fileTable;
////////////    private static JTextArea statsTextArea = new JTextArea();
////////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////////
////////////    public static void main(String[] args) {
////////////        // 创建并显示主窗口
////////////        frame = createMainFrame();
////////////        frame.setVisible(true);
////////////
////////////        // 加载配置文件
////////////        loadConfigFile(CONFIG_FILE);
////////////    }
////////////
////////////    private static void loadConfigFile(String configFile) {
////////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////////
////////////            List<XWPFTable> configTables = configDoc.getTables();
////////////
////////////            for (XWPFTable table : configTables) {
////////////                for (XWPFTableRow row : table.getRows()) {
////////////                    if (row.getTableCells().size() >= 2) {
////////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////////                        String key = labelCell.getText().trim();
////////////                        String value = valueCell.getText().trim();
////////////                        configMap.put(key, value);
////////////                        configTableModel.addRow(new Object[]{key, value});
////////////                        keyComboBox.addItem(key);
////////////                    }
////////////                }
////////////            }
////////////        } catch (IOException e) {
////////////            e.printStackTrace();
////////////        }
////////////    }
////////////
////////////    private static void saveConfigFile(String configFile) {
////////////        try (FileOutputStream fos = new FileOutputStream(configFile);
////////////             XWPFDocument configDoc = new XWPFDocument()) {
////////////
////////////            XWPFTable table = configDoc.createTable();
////////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////                XWPFTableRow row = table.createRow();
////////////                row.getCell(0).setText(entry.getKey());
////////////                row.getCell(1).setText(entry.getValue());
////////////            }
////////////            configDoc.write(fos);
////////////        } catch (IOException e) {
////////////            e.printStackTrace();
////////////        }
////////////    }
////////////
////////////    private static String convertDocToDocx(String sourceFile) {
////////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////////        try {
////////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////////            doc.save(docxFile, SaveFormat.DOCX);
////////////        } catch (Exception e) {
////////////            throw new RuntimeException(e);
////////////        }
////////////        return docxFile;
////////////    }
////////////
////////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////////
////////////            StringBuilder originalContent = new StringBuilder();
////////////            StringBuilder modifiedContent = new StringBuilder();
////////////
////////////            List<XWPFTable> tables = document.getTables();
////////////
////////////            for (XWPFTable table : tables) {
////////////                for (XWPFTableRow row : table.getRows()) {
////////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////////                        String text = cell.getText().replaceAll("\\s+", "");
////////////                        originalContent.append(text).append(" ");
////////////
////////////                        if (configMap.containsKey(text)) {
////////////                            String newValue = configMap.get(text);
////////////                            modifiedContent.append(newValue).append(" ");
////////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////////                                nextCell.removeParagraph(0);
////////////                                XWPFParagraph p = nextCell.addParagraph();
////////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////////                                XWPFRun r = p.createRun();
////////////                                r.setText(newValue);
////////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////////                            }
////////////                        }
////////////                    }
////////////                }
////////////            }
////////////
////////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////////                List<XWPFRun> runs = paragraph.getRuns();
////////////                for (int i = 0; i < runs.size(); i++) {
////////////                    XWPFRun run = runs.get(i);
////////////                    String text = run.getText(0);
////////////                    if (text != null) {
////////////                        originalContent.append(text).append(" ");
////////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////////                            String key = entry.getKey();
////////////                            String value = entry.getValue();
////////////                            if (text.trim().equals(key)) {
////////////                                int j = i + 1;
////////////                                while (j < runs.size()) {
////////////                                    XWPFRun nextRun = runs.get(j);
////////////                                    String nextText = nextRun.getText(0);
////////////                                    if (nextText != null && !nextText.contains(":")) {
////////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////////                                        newRun.setText(value);
////////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////////                                        newRun.setFontFamily("仿宋_GB2312");
////////////                                        newRun.setFontSize(14);
////////////                                        paragraph.removeRun(j);
////////////                                        break;
////////////                                    }
////////////                                    j++;
////////////                                }
////////////                                i = j;
////////////                                break;
////////////                            }
////////////                        }
////////////                    }
////////////                }
////////////            }
////////////
////////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////////            document.write(fos);
////////////        }
////////////    }
////////////
////////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////////        progressBar.setValue(progress);
////////////    }
////////////
////////////    private static JFrame createMainFrame() {
////////////        JFrame frame = new JFrame("文档处理工具");
////////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////////        frame.setSize(1200, 800);
////////////
////////////        JPanel panel = new JPanel(new BorderLayout());
////////////        frame.add(panel);
////////////
////////////        JPanel configPanel = new JPanel(new BorderLayout());
////////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////////        configTable = new JTable(configTableModel);
////////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////////
////////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////////        configInputPanel.add(new JLabel("Key:"));
////////////        configInputPanel.add(keyComboBox);
////////////        configInputPanel.add(new JLabel("Value:"));
////////////        configInputPanel.add(valueField);
////////////
////////////        JButton addButton = new JButton("添加/更新");
////////////        addButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                String key = (String) keyComboBox.getSelectedItem();
////////////                String value = valueField.getText().trim();
////////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
////////////                    configMap.put(key, value);
////////////                    boolean keyExists = false;
////////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////////                            configTableModel.setValueAt(value, i, 1);
////////////                            keyExists = true;
////////////                            break;
////////////                        }
////////////                    }
////////////                    if (!keyExists) {
////////////                        configTableModel.addRow(new Object[]{key, value});
////////////                        keyComboBox.addItem(key);
////////////                    }
////////////                    saveConfigFile(CONFIG_FILE);
////////////                }
////////////            }
////////////        });
////////////        configInputPanel.add(addButton);
////////////
////////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////////
////////////        JPanel filePanel = new JPanel(new BorderLayout());
////////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////////        fileTable = new JTable(fileTableModel);
////////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
////////////            if (!event.getValueIsAdjusting()) {
////////////                int selectedRow = fileTable.getSelectedRow();
////////////                if (selectedRow != -1) {
////////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////////                    displayFilePreview(fileName, originalContent, modifiedContent);
////////////                }
////////////            }
////////////        });
////////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////////
////////////        JButton refreshButton = new JButton("刷新预览");
////////////        refreshButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                fileTableModel.setRowCount(0);
////////////                processFiles();
////////////            }
////////////        });
////////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
////////////
////////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////////
////////////        JButton startButton = new JButton("开始执行");
////////////        startButton.addActionListener(new ActionListener() {
////////////            @Override
////////////            public void actionPerformed(ActionEvent e) {
////////////                processFiles();
////////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
////////////            }
////////////        });
////////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////////
////////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
////////////        statsTextArea.setEditable(false);
////////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////////
////////////        panel.add(configPanel, BorderLayout.NORTH);
////////////        panel.add(filePanel, BorderLayout.CENTER);
////////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////////        panel.add(statsPanel, BorderLayout.EAST);
////////////
////////////        return frame;
////////////    }
////////////
////////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////////        previewFrame.setSize(600, 400);
////////////        JTextArea previewTextArea = new JTextArea();
////////////        previewTextArea.setEditable(false);
////////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////////        previewFrame.add(new JScrollPane(previewTextArea));
////////////        previewFrame.setVisible(true);
////////////    }
////////////
////////////    private static void processFiles() {
////////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////////        File[] listOfFiles = folder.listFiles();
////////////        if (listOfFiles == null) {
////////////            return;
////////////        }
////////////        int totalFiles = listOfFiles.length;
////////////        int processedFiles = 0;
////////////        long startTime = System.currentTimeMillis();
////////////        for (File file : listOfFiles) {
////////////            if (file.isFile()) {
////////////                String sourceFile = file.getAbsolutePath();
////////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////////                if (sourceFile.endsWith(".doc")) {
////////////                    sourceFile = convertDocToDocx(sourceFile);
////////////                }
////////////                if (!sourceFile.endsWith(".docx")) {
////////////                    continue;
////////////                }
////////////                try {
////////////                    modifyDocument(sourceFile, outputFile);
////////////                    processedFiles++;
////////////                    updateProgress(processedFiles, totalFiles);
////////////                } catch (IOException e) {
////////////                    e.printStackTrace();
////////////                }
////////////            }
////////////        }
////////////        long endTime = System.currentTimeMillis();
////////////        long duration = endTime - startTime;
////////////        displayStats(totalFiles, processedFiles, duration);
////////////    }
////////////
////////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
////////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
////////////        statsTextArea.setText(stats);
////////////    }
////////////}
//////////
//////////
//////////
//////////package org.example;
//////////
//////////import com.aspose.words.SaveFormat;
//////////import org.apache.poi.xwpf.usermodel.*;
//////////
//////////import javax.swing.*;
//////////import javax.swing.table.DefaultTableModel;
//////////import java.awt.*;
//////////import java.awt.event.ActionEvent;
//////////import java.awt.event.ActionListener;
//////////import java.io.File;
//////////import java.io.FileInputStream;
//////////import java.io.FileOutputStream;
//////////import java.io.IOException;
//////////import java.util.HashMap;
//////////import java.util.List;
//////////import java.util.Map;
//////////
//////////public class WordModifier {
//////////    private static Map<String, String> configMap = new HashMap<>();
//////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////////    private static JFrame frame;
//////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////////    private static JTextField valueField = new JTextField();
//////////    private static JTable configTable;
//////////    private static JTable fileTable;
//////////    private static JTextArea statsTextArea = new JTextArea();
//////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////////
//////////    public static void main(String[] args) {
//////////        // 创建并显示主窗口
//////////        frame = createMainFrame();
//////////        frame.setVisible(true);
//////////
//////////        // 加载配置文件
//////////        loadConfigFile(CONFIG_FILE);
//////////    }
//////////
//////////    private static void loadConfigFile(String configFile) {
//////////        try (FileInputStream configFis = new FileInputStream(configFile);
//////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////////
//////////            List<XWPFTable> configTables = configDoc.getTables();
//////////
//////////            for (XWPFTable table : configTables) {
//////////                for (XWPFTableRow row : table.getRows()) {
//////////                    if (row.getTableCells().size() >= 2) {
//////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////////                        String key = labelCell.getText().trim();
//////////                        String value = valueCell.getText().trim();
//////////                        configMap.put(key, value);
//////////                        configTableModel.addRow(new Object[]{key, value});
//////////                        keyComboBox.addItem(key);
//////////                    }
//////////                }
//////////            }
//////////        } catch (IOException e) {
//////////            e.printStackTrace();
//////////        }
//////////    }
//////////
//////////    private static void saveConfigFile(String configFile) {
//////////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////////             XWPFDocument configDoc = new XWPFDocument()) {
//////////
//////////            XWPFTable table = configDoc.createTable();
//////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////                XWPFTableRow row = table.createRow();
//////////                row.getCell(0).setText(entry.getKey());
//////////                row.getCell(1).setText(entry.getValue());
//////////            }
//////////            configDoc.write(fos);
//////////        } catch (IOException e) {
//////////            e.printStackTrace();
//////////        }
//////////    }
//////////
//////////    private static String convertDocToDocx(String sourceFile) {
//////////        String docxFile = sourceFile.replace(".doc", ".docx");
//////////        try {
//////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////////            doc.save(docxFile, SaveFormat.DOCX);
//////////        } catch (Exception e) {
//////////            throw new RuntimeException(e);
//////////        }
//////////        return docxFile;
//////////    }
//////////
//////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////////             FileOutputStream fos = new FileOutputStream(outputFile);
//////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////////
//////////            StringBuilder originalContent = new StringBuilder();
//////////            StringBuilder modifiedContent = new StringBuilder();
//////////
//////////            List<XWPFTable> tables = document.getTables();
//////////
//////////            for (XWPFTable table : tables) {
//////////                for (XWPFTableRow row : table.getRows()) {
//////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////////                        XWPFTableCell cell = row.getTableCells().get(i);
//////////                        String text = cell.getText().replaceAll("\\s+", "");
//////////                        originalContent.append(text).append(" ");
//////////
//////////                        if (configMap.containsKey(text)) {
//////////                            String newValue = configMap.get(text);
//////////                            modifiedContent.append(newValue).append(" ");
//////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////////                                nextCell.removeParagraph(0);
//////////                                XWPFParagraph p = nextCell.addParagraph();
//////////                                p.setAlignment(ParagraphAlignment.CENTER);
//////////                                XWPFRun r = p.createRun();
//////////                                r.setText(newValue);
//////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////////                            }
//////////                        }
//////////                    }
//////////                }
//////////            }
//////////
//////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////////                List<XWPFRun> runs = paragraph.getRuns();
//////////                for (int i = 0; i < runs.size(); i++) {
//////////                    XWPFRun run = runs.get(i);
//////////                    String text = run.getText(0);
//////////                    if (text != null) {
//////////                        originalContent.append(text).append(" ");
//////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////////                            String key = entry.getKey();
//////////                            String value = entry.getValue();
//////////                            if (text.trim().equals(key)) {
//////////                                int j = i + 1;
//////////                                while (j < runs.size()) {
//////////                                    XWPFRun nextRun = runs.get(j);
//////////                                    String nextText = nextRun.getText(0);
//////////                                    if (nextText != null && !nextText.contains(":")) {
//////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////////                                        newRun.setText(value);
//////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////////                                        newRun.setFontFamily("仿宋_GB2312");
//////////                                        newRun.setFontSize(14);
//////////                                        paragraph.removeRun(j);
//////////                                        break;
//////////                                    }
//////////                                    j++;
//////////                                }
//////////                                i = j;
//////////                                break;
//////////                            }
//////////                        }
//////////                    }
//////////                }
//////////            }
//////////
//////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////////            document.write(fos);
//////////        }
//////////    }
//////////
//////////    private static void updateProgress(int processedFiles, int totalFiles) {
//////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////////        progressBar.setValue(progress);
//////////    }
//////////
//////////    private static JFrame createMainFrame() {
//////////        JFrame frame = new JFrame("文档处理工具");
//////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////////        frame.setSize(1200, 800);
//////////
//////////        JPanel panel = new JPanel(new BorderLayout());
//////////        frame.add(panel);
//////////
//////////        JPanel configPanel = new JPanel(new BorderLayout());
//////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////////        configTable = new JTable(configTableModel);
//////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////////
//////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////////        configInputPanel.add(new JLabel("Key:"));
//////////        configInputPanel.add(keyComboBox);
//////////        configInputPanel.add(new JLabel("Value:"));
//////////        configInputPanel.add(valueField);
//////////
//////////        JButton addButton = new JButton("添加/更新");
//////////        addButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                String key = (String) keyComboBox.getSelectedItem();
//////////                String value = valueField.getText().trim();
//////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
//////////                    configMap.put(key, value);
//////////                    boolean keyExists = false;
//////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////////                            configTableModel.setValueAt(value, i, 1);
//////////                            keyExists = true;
//////////                            break;
//////////                        }
//////////                    }
//////////                    if (!keyExists) {
//////////                        configTableModel.addRow(new Object[]{key, value});
//////////                        keyComboBox.addItem(key);
//////////                    }
//////////                    saveConfigFile(CONFIG_FILE);
//////////                }
//////////            }
//////////        });
//////////        configInputPanel.add(addButton);
//////////
//////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////////
//////////        JPanel filePanel = new JPanel(new BorderLayout());
//////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////////        fileTable = new JTable(fileTableModel);
//////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
//////////            if (!event.getValueIsAdjusting()) {
//////////                int selectedRow = fileTable.getSelectedRow();
//////////                if (selectedRow != -1) {
//////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////////                    displayFilePreview(fileName, originalContent, modifiedContent);
//////////                }
//////////            }
//////////        });
//////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////////
//////////        JButton refreshButton = new JButton("刷新预览");
//////////        refreshButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                fileTableModel.setRowCount(0);
//////////                processFiles();
//////////            }
//////////        });
//////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////////
//////////        JPanel progressPanel = new JPanel(new BorderLayout());
//////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////////
//////////        JButton startButton = new JButton("开始执行");
//////////        startButton.addActionListener(new ActionListener() {
//////////            @Override
//////////            public void actionPerformed(ActionEvent e) {
//////////                processFiles();
//////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////////            }
//////////        });
//////////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////////
//////////        JPanel statsPanel = new JPanel(new BorderLayout());
//////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////////        statsTextArea.setEditable(false);
//////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////////
//////////        panel.add(configPanel, BorderLayout.NORTH);
//////////        panel.add(filePanel, BorderLayout.CENTER);
//////////        panel.add(progressPanel, BorderLayout.SOUTH);
//////////        panel.add(statsPanel, BorderLayout.EAST);
//////////
//////////        return frame;
//////////    }
//////////
//////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////////        previewFrame.setSize(600, 400);
//////////        JTextArea previewTextArea = new JTextArea();
//////////        previewTextArea.setEditable(false);
//////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////////        previewFrame.add(new JScrollPane(previewTextArea));
//////////        previewFrame.setVisible(true);
//////////    }
//////////
//////////    private static void processFiles() {
//////////        File folder = new File("Z:\\Desktop\\测试\\in");
//////////        File[] listOfFiles = folder.listFiles();
//////////        if (listOfFiles == null) {
//////////            return;
//////////        }
//////////        int totalFiles = listOfFiles.length;
//////////        int processedFiles = 0;
//////////        long startTime = System.currentTimeMillis();
//////////        for (File file : listOfFiles) {
//////////            if (file.isFile()) {
//////////                String sourceFile = file.getAbsolutePath();
//////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////////                if (sourceFile.endsWith(".doc")) {
//////////                    sourceFile = convertDocToDocx(sourceFile);
//////////                }
//////////                if (!sourceFile.endsWith(".docx")) {
//////////                    continue;
//////////                }
//////////                try {
//////////                    modifyDocument(sourceFile, outputFile);
//////////                    processedFiles++;
//////////                    updateProgress(processedFiles, totalFiles);
//////////                } catch (IOException e) {
//////////                    e.printStackTrace();
//////////                }
//////////            }
//////////        }
//////////        long endTime = System.currentTimeMillis();
//////////        long duration = endTime - startTime;
//////////        displayStats(totalFiles, processedFiles, duration);
//////////    }
//////////
//////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////////        statsTextArea.setText(stats);
//////////    }
//////////}
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////
////////package org.example;
////////
////////import com.aspose.words.SaveFormat;
////////import org.apache.poi.xwpf.usermodel.*;
////////
////////import javax.swing.*;
////////import javax.swing.table.DefaultTableModel;
////////import java.awt.*;
////////import java.awt.event.ActionEvent;
////////import java.awt.event.ActionListener;
////////import java.io.File;
////////import java.io.FileInputStream;
////////import java.io.FileOutputStream;
////////import java.io.IOException;
////////import java.util.HashMap;
////////import java.util.List;
////////import java.util.Map;
////////
////////public class WordModifier {
////////    private static Map<String, String> configMap = new HashMap<>();
////////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
////////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
////////    private static JProgressBar progressBar = new JProgressBar(0, 100);
////////    private static JFrame frame;
////////    private static JComboBox<String> keyComboBox = new JComboBox<>();
////////    private static JTextField valueField = new JTextField();
////////    private static JTable configTable;
////////    private static JTable fileTable;
////////    private static JTextArea statsTextArea = new JTextArea();
////////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
////////
////////    public static void main(String[] args) {
////////        // 创建并显示主窗口
////////        frame = createMainFrame();
////////        frame.setVisible(true);
////////
////////        // 加载配置文件
////////        loadConfigFile(CONFIG_FILE);
////////    }
////////
////////    private static void loadConfigFile(String configFile) {
////////        try (FileInputStream configFis = new FileInputStream(configFile);
////////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
////////
////////            List<XWPFTable> configTables = configDoc.getTables();
////////
////////            for (XWPFTable table : configTables) {
////////                for (XWPFTableRow row : table.getRows()) {
////////                    if (row.getTableCells().size() >= 2) {
////////                        XWPFTableCell labelCell = row.getTableCells().get(0);
////////                        XWPFTableCell valueCell = row.getTableCells().get(1);
////////                        String key = labelCell.getText().trim();
////////                        String value = valueCell.getText().trim();
////////                        configMap.put(key, value);
////////                        configTableModel.addRow(new Object[]{key, value});
////////                        keyComboBox.addItem(key);
////////                    }
////////                }
////////            }
////////        } catch (IOException e) {
////////            e.printStackTrace();
////////        }
////////    }
////////
////////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
////////        try (FileOutputStream fos = new FileOutputStream(configFile);
////////             XWPFDocument configDoc = new XWPFDocument()) {
////////
////////            XWPFTable table = configDoc.createTable();
////////            boolean firstRow = true;
////////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////                XWPFTableRow row;
////////                if (firstRow) {
////////                    row = table.getRow(0); // 使用已经存在的第一行
////////                    firstRow = false;
////////                } else {
////////                    row = table.createRow();
////////                }
////////                row.getCell(0).setText(entry.getKey());
////////                row.getCell(1).setText(entry.getValue());
////////            }
////////            configDoc.write(fos);
////////        } catch (IOException e) {
////////            // 这里可以使用更合适的异常处理方式
////////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
////////        }
////////    }
////////
////////
////////    private static String convertDocToDocx(String sourceFile) {
////////        String docxFile = sourceFile.replace(".doc", ".docx");
////////        try {
////////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
////////            doc.save(docxFile, SaveFormat.DOCX);
////////        } catch (Exception e) {
////////            throw new RuntimeException(e);
////////        }
////////        return docxFile;
////////    }
////////
////////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
////////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
////////             FileOutputStream fos = new FileOutputStream(outputFile);
////////             XWPFDocument document = new XWPFDocument(sourceFis)) {
////////
////////            StringBuilder originalContent = new StringBuilder();
////////            StringBuilder modifiedContent = new StringBuilder();
////////
////////            List<XWPFTable> tables = document.getTables();
////////
////////            for (XWPFTable table : tables) {
////////                for (XWPFTableRow row : table.getRows()) {
////////                    for (int i = 0; i < row.getTableCells().size(); i++) {
////////                        XWPFTableCell cell = row.getTableCells().get(i);
////////                        String text = cell.getText().replaceAll("\\s+", "");
////////                        originalContent.append(text).append(" ");
////////
////////                        if (configMap.containsKey(text)) {
////////                            String newValue = configMap.get(text);
////////                            modifiedContent.append(newValue).append(" ");
////////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
////////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
////////                                nextCell.removeParagraph(0);
////////                                XWPFParagraph p = nextCell.addParagraph();
////////                                p.setAlignment(ParagraphAlignment.CENTER);
////////                                XWPFRun r = p.createRun();
////////                                r.setText(newValue);
////////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
////////                            }
////////                        }
////////                    }
////////                }
////////            }
////////
////////            for (XWPFParagraph paragraph : document.getParagraphs()) {
////////                List<XWPFRun> runs = paragraph.getRuns();
////////                for (int i = 0; i < runs.size(); i++) {
////////                    XWPFRun run = runs.get(i);
////////                    String text = run.getText(0);
////////                    if (text != null) {
////////                        originalContent.append(text).append(" ");
////////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
////////                            String key = entry.getKey();
////////                            String value = entry.getValue();
////////                            if (text.trim().equals(key)) {
////////                                int j = i + 1;
////////                                while (j < runs.size()) {
////////                                    XWPFRun nextRun = runs.get(j);
////////                                    String nextText = nextRun.getText(0);
////////                                    if (nextText != null && !nextText.contains(":")) {
////////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
////////                                        newRun.setText(value);
////////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
////////                                        newRun.setFontFamily("仿宋_GB2312");
////////                                        newRun.setFontSize(14);
////////                                        paragraph.removeRun(j);
////////                                        break;
////////                                    }
////////                                    j++;
////////                                }
////////                                i = j;
////////                                break;
////////                            }
////////                        }
////////                    }
////////                }
////////            }
////////
////////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
////////            document.write(fos);
////////        }
////////    }
////////
////////    private static void updateProgress(int processedFiles, int totalFiles) {
////////        int progress = (int) ((double) processedFiles / totalFiles * 100);
////////        progressBar.setValue(progress);
////////    }
////////
////////    private static JFrame createMainFrame() {
////////        JFrame frame = new JFrame("文档处理工具");
////////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
////////        frame.setSize(1200, 800);
////////
////////        JPanel panel = new JPanel(new BorderLayout());
////////        frame.add(panel);
////////
////////        JPanel configPanel = new JPanel(new BorderLayout());
////////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
////////        configTable = new JTable(configTableModel);
////////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
////////
////////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
////////        configInputPanel.add(new JLabel("Key:"));
////////        configInputPanel.add(keyComboBox);
////////        configInputPanel.add(new JLabel("Value:"));
////////        configInputPanel.add(valueField);
////////
////////        JButton addButton = new JButton("添加/更新");
////////        addButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                String key = (String) keyComboBox.getSelectedItem();
////////                String value = valueField.getText().trim();
////////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
////////                    configMap.put(key, value);
////////                    boolean keyExists = false;
////////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
////////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
////////                            configTableModel.setValueAt(value, i, 1);
////////                            keyExists = true;
////////                            break;
////////                        }
////////                    }
////////                    if (!keyExists) {
////////                        configTableModel.addRow(new Object[]{key, value});
////////                        keyComboBox.addItem(key);
////////                    }
////////                    saveConfigFile(CONFIG_FILE,configMap);
////////                }
////////            }
////////        });
////////        configInputPanel.add(addButton);
////////
////////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
////////
////////        JPanel filePanel = new JPanel(new BorderLayout());
////////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
////////        fileTable = new JTable(fileTableModel);
////////        fileTable.getSelectionModel().addListSelectionListener(event -> {
////////            if (!event.getValueIsAdjusting()) {
////////                int selectedRow = fileTable.getSelectedRow();
////////                if (selectedRow != -1) {
////////                    String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
////////                    String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
////////                    String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
////////                    displayFilePreview(fileName, originalContent, modifiedContent);
////////                }
////////            }
////////        });
////////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
////////
////////        JButton refreshButton = new JButton("刷新预览");
////////        refreshButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                fileTableModel.setRowCount(0);
////////                processFiles();
////////            }
////////        });
////////        filePanel.add(refreshButton, BorderLayout.SOUTH);
////////
////////        JPanel progressPanel = new JPanel(new BorderLayout());
////////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
////////        progressPanel.add(progressBar, BorderLayout.CENTER);
////////
////////        JButton startButton = new JButton("开始执行");
////////        startButton.addActionListener(new ActionListener() {
////////            @Override
////////            public void actionPerformed(ActionEvent e) {
////////                processFiles();
////////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
////////            }
////////        });
////////        progressPanel.add(startButton, BorderLayout.SOUTH);
////////
////////        JPanel statsPanel = new JPanel(new BorderLayout());
////////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
////////        statsTextArea.setEditable(false);
////////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
////////
////////        panel.add(configPanel, BorderLayout.NORTH);
////////        panel.add(filePanel, BorderLayout.CENTER);
////////        panel.add(progressPanel, BorderLayout.SOUTH);
////////        panel.add(statsPanel, BorderLayout.EAST);
////////
////////        return frame;
////////    }
////////
////////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
////////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
////////        previewFrame.setSize(600, 400);
////////        JTextArea previewTextArea = new JTextArea();
////////        previewTextArea.setEditable(false);
////////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
////////        previewFrame.add(new JScrollPane(previewTextArea));
////////        previewFrame.setVisible(true);
////////    }
////////
////////    private static void processFiles() {
////////        File folder = new File("Z:\\Desktop\\测试\\in");
////////        File[] listOfFiles = folder.listFiles();
////////        if (listOfFiles == null) {
////////            return;
////////        }
////////        int totalFiles = listOfFiles.length;
////////        int processedFiles = 0;
////////        long startTime = System.currentTimeMillis();
////////        for (File file : listOfFiles) {
////////            if (file.isFile()) {
////////                String sourceFile = file.getAbsolutePath();
////////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
////////                if (sourceFile.endsWith(".doc")) {
////////                    sourceFile = convertDocToDocx(sourceFile);
////////                }
////////                if (!sourceFile.endsWith(".docx")) {
////////                    continue;
////////                }
////////                try {
////////                    modifyDocument(sourceFile, outputFile);
////////                    processedFiles++;
////////                    updateProgress(processedFiles, totalFiles);
////////                } catch (IOException e) {
////////                    e.printStackTrace();
////////                }
////////            }
////////        }
////////        long endTime = System.currentTimeMillis();
////////        long duration = endTime - startTime;
////////        displayStats(totalFiles, processedFiles, duration);
////////    }
////////
////////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
////////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
////////        statsTextArea.setText(stats);
////////    }
////////}
////////
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////
//////package org.example;
//////
//////import com.aspose.words.SaveFormat;
//////import org.apache.poi.xwpf.usermodel.*;
//////
//////import javax.swing.*;
//////import javax.swing.event.ListSelectionEvent;
//////import javax.swing.event.ListSelectionListener;
//////import javax.swing.table.DefaultTableModel;
//////import java.awt.*;
//////import java.awt.event.ActionEvent;
//////import java.awt.event.ActionListener;
//////import java.io.File;
//////import java.io.FileInputStream;
//////import java.io.FileOutputStream;
//////import java.io.IOException;
//////import java.util.HashMap;
//////import java.util.LinkedHashMap;
//////import java.util.List;
//////import java.util.Map;
//////
//////public class WordModifier {
//////    private static Map<String, String> configMap = new LinkedHashMap<>();
//////
//////    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//////    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//////    private static JProgressBar progressBar = new JProgressBar(0, 100);
//////    private static JFrame frame;
//////    private static JComboBox<String> keyComboBox = new JComboBox<>();
//////    private static JTextField valueField = new JTextField();
//////    private static JTable configTable;
//////    private static JTable fileTable;
//////    private static JTextArea statsTextArea = new JTextArea();
//////    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//////
//////    public static void main(String[] args) {
//////        // 创建并显示主窗口
//////        frame = createMainFrame();
//////        frame.setVisible(true);
//////
//////        // 加载配置文件
//////        loadConfigFile(CONFIG_FILE);
//////    }
//////
//////    private static void loadConfigFile(String configFile) {
//////        try (FileInputStream configFis = new FileInputStream(configFile);
//////             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//////
//////            List<XWPFTable> configTables = configDoc.getTables();
//////
//////            for (XWPFTable table : configTables) {
//////                for (XWPFTableRow row : table.getRows()) {
//////                    if (row.getTableCells().size() >= 2) {
//////                        XWPFTableCell labelCell = row.getTableCells().get(0);
//////                        XWPFTableCell valueCell = row.getTableCells().get(1);
//////                        String key = labelCell.getText().trim();
//////                        String value = valueCell.getText().trim();
//////                        configMap.put(key, value);
//////                        configTableModel.addRow(new Object[]{key, value});
//////                        keyComboBox.addItem(key);
//////                    }
//////                }
//////            }
//////        } catch (IOException e) {
//////            e.printStackTrace();
//////        }
//////    }
//////
//////    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//////        try (FileOutputStream fos = new FileOutputStream(configFile);
//////             XWPFDocument configDoc = new XWPFDocument()) {
//////
//////            XWPFTable table = configDoc.createTable();
//////            boolean firstRow = true;
//////            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////                XWPFTableRow row;
//////                if (firstRow) {
//////                    row = table.getRow(0); // 使用已经存在的第一行
//////                    firstRow = false;
//////                } else {
//////                    row = table.createRow();
//////                }
//////                XWPFTableCell keyCell = row.getCell(0);
//////                if (keyCell == null) {
//////                    keyCell = row.addNewTableCell();
//////                }
//////                keyCell.setText(entry.getKey());
//////
//////                XWPFTableCell valueCell;
//////                if (row.getTableCells().size() > 1) {
//////                    valueCell = row.getCell(1);
//////                } else {
//////                    valueCell = row.addNewTableCell();
//////                }
//////                valueCell.setText(entry.getValue());
//////            }
//////            configDoc.write(fos);
//////        } catch (IOException e) {
//////            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//////        }
//////    }
//////
//////
//////    private static String convertDocToDocx(String sourceFile) {
//////        String docxFile = sourceFile.replace(".doc", ".docx");
//////        try {
//////            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//////            doc.save(docxFile, SaveFormat.DOCX);
//////        } catch (Exception e) {
//////            throw new RuntimeException(e);
//////        }
//////        return docxFile;
//////    }
//////
//////    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//////        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//////             FileOutputStream fos = new FileOutputStream(outputFile);
//////             XWPFDocument document = new XWPFDocument(sourceFis)) {
//////
//////            StringBuilder originalContent = new StringBuilder();
//////            StringBuilder modifiedContent = new StringBuilder();
//////
//////            List<XWPFTable> tables = document.getTables();
//////
//////            for (XWPFTable table : tables) {
//////                for (XWPFTableRow row : table.getRows()) {
//////                    for (int i = 0; i < row.getTableCells().size(); i++) {
//////                        XWPFTableCell cell = row.getTableCells().get(i);
//////                        String text = cell.getText().replaceAll("\\s+", "");
//////                        originalContent.append(text).append(" ");
//////
//////                        if (configMap.containsKey(text)) {
//////                            String newValue = configMap.get(text);
//////                            modifiedContent.append(newValue).append(" ");
//////                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//////                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//////                                nextCell.removeParagraph(0);
//////                                XWPFParagraph p = nextCell.addParagraph();
//////                                p.setAlignment(ParagraphAlignment.CENTER);
//////                                XWPFRun r = p.createRun();
//////                                r.setText(newValue);
//////                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//////                            }
//////                        }
//////                    }
//////                }
//////            }
//////
//////            for (XWPFParagraph paragraph : document.getParagraphs()) {
//////                List<XWPFRun> runs = paragraph.getRuns();
//////                for (int i = 0; i < runs.size(); i++) {
//////                    XWPFRun run = runs.get(i);
//////                    String text = run.getText(0);
//////                    if (text != null) {
//////                        originalContent.append(text).append(" ");
//////                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//////                            String key = entry.getKey();
//////                            String value = entry.getValue();
//////                            if (text.trim().equals(key)) {
//////                                int j = i + 1;
//////                                while (j < runs.size()) {
//////                                    XWPFRun nextRun = runs.get(j);
//////                                    String nextText = nextRun.getText(0);
//////                                    if (nextText != null && !nextText.contains(":")) {
//////                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//////                                        newRun.setText(value);
//////                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//////                                        newRun.setFontFamily("仿宋_GB2312");
//////                                        newRun.setFontSize(14);
//////                                        paragraph.removeRun(j);
//////                                        break;
//////                                    }
//////                                    j++;
//////                                }
//////                                i = j;
//////                                break;
//////                            }
//////                        }
//////                    }
//////                }
//////            }
//////
//////            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//////            document.write(fos);
//////        }
//////    }
//////
//////    private static void updateProgress(int processedFiles, int totalFiles) {
//////        int progress = (int) ((double) processedFiles / totalFiles * 100);
//////        progressBar.setValue(progress);
//////    }
//////
//////    private static JFrame createMainFrame() {
//////        JFrame frame = new JFrame("文档处理工具");
//////        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//////        frame.setSize(1200, 800);
//////
//////        JPanel panel = new JPanel(new BorderLayout());
//////        frame.add(panel);
//////
//////        JPanel configPanel = new JPanel(new BorderLayout());
//////        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
//////        configTable = new JTable(configTableModel);
//////        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//////
//////        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
//////        configInputPanel.add(new JLabel("Key:"));
//////        configInputPanel.add(keyComboBox);
//////        configInputPanel.add(new JLabel("Value:"));
//////        configInputPanel.add(valueField);
//////
//////        JButton addButton = new JButton("添加/更新");
//////        addButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                String key = (String) keyComboBox.getSelectedItem();
//////                String value = valueField.getText().trim();
//////                if (key != null && !key.isEmpty() && !value.isEmpty()) {
//////                    configMap.put(key, value);
//////                    boolean keyExists = false;
//////                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
//////                        if (configTableModel.getValueAt(i, 0).equals(key)) {
//////                            configTableModel.setValueAt(value, i, 1);
//////                            keyExists = true;
//////                            break;
//////                        }
//////                    }
//////                    if (!keyExists) {
//////                        configTableModel.addRow(new Object[]{key, value});
//////                        keyComboBox.addItem(key);
//////                    }
//////                    saveConfigFile(CONFIG_FILE, configMap);
//////                }
//////            }
//////        });
//////        configInputPanel.add(addButton);
//////
//////        configPanel.add(configInputPanel, BorderLayout.SOUTH);
//////
//////        JPanel filePanel = new JPanel(new BorderLayout());
//////        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
//////        fileTable = new JTable(fileTableModel);
//////        fileTable.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
//////            @Override
//////            public void valueChanged(ListSelectionEvent event) {
//////                if (!event.getValueIsAdjusting()) {
//////                    int selectedRow = fileTable.getSelectedRow();
//////                    if (selectedRow != -1) {
//////                        String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
//////                        String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
//////                        String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
//////                        displayFilePreview(fileName, originalContent, modifiedContent);
//////                    }
//////                }
//////            }
//////        });
//////        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//////
//////        JButton refreshButton = new JButton("刷新预览");
//////        refreshButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                fileTableModel.setRowCount(0);
//////                processFiles();
//////            }
//////        });
//////        filePanel.add(refreshButton, BorderLayout.SOUTH);
//////
//////        JPanel progressPanel = new JPanel(new BorderLayout());
//////        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
//////        progressPanel.add(progressBar, BorderLayout.CENTER);
//////
//////        JButton startButton = new JButton("开始执行");
//////        startButton.addActionListener(new ActionListener() {
//////            @Override
//////            public void actionPerformed(ActionEvent e) {
//////                processFiles();
//////                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
//////            }
//////        });
//////        progressPanel.add(startButton, BorderLayout.SOUTH);
//////
//////        JPanel statsPanel = new JPanel(new BorderLayout());
//////        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
//////        statsTextArea.setEditable(false);
//////        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//////
//////        panel.add(configPanel, BorderLayout.NORTH);
//////        panel.add(filePanel, BorderLayout.CENTER);
//////        panel.add(progressPanel, BorderLayout.SOUTH);
//////        panel.add(statsPanel, BorderLayout.EAST);
//////
//////        return frame;
//////    }
//////
//////    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
//////        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
//////        previewFrame.setSize(600, 400);
//////        JTextArea previewTextArea = new JTextArea();
//////        previewTextArea.setEditable(false);
//////        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
//////        previewFrame.add(new JScrollPane(previewTextArea));
//////        previewFrame.setVisible(true);
//////    }
//////
//////    private static void processFiles() {
//////        File folder = new File("Z:\\Desktop\\测试\\in");
//////        File[] listOfFiles = folder.listFiles();
//////        if (listOfFiles == null) {
//////            return;
//////        }
//////        int totalFiles = listOfFiles.length;
//////        int processedFiles = 0;
//////        long startTime = System.currentTimeMillis();
//////        for (File file : listOfFiles) {
//////            if (file.isFile()) {
//////                String sourceFile = file.getAbsolutePath();
//////                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
//////                if (sourceFile.endsWith(".doc")) {
//////                    sourceFile = convertDocToDocx(sourceFile);
//////                }
//////                if (!sourceFile.endsWith(".docx")) {
//////                    continue;
//////                }
//////                try {
//////                    modifyDocument(sourceFile, outputFile);
//////                    processedFiles++;
//////                    updateProgress(processedFiles, totalFiles);
//////                } catch (IOException e) {
//////                    e.printStackTrace();
//////                }
//////            }
//////        }
//////        long endTime = System.currentTimeMillis();
//////        long duration = endTime - startTime;
//////        displayStats(totalFiles, processedFiles, duration);
//////    }
//////
//////    private static void displayStats(int totalFiles, int processedFiles, long duration) {
//////        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
//////        statsTextArea.setText(stats);
//////    }
//////}
////
////
////
////
////
////
////
////

package org.example;

import com.aspose.words.SaveFormat;
import org.apache.poi.xwpf.usermodel.*;

import javax.swing.*;
import javax.swing.event.ListSelectionEvent;
import javax.swing.event.ListSelectionListener;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.List;

public class WordModifier {
    private static Map<String, String> configMap = new LinkedHashMap<>();

    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();


    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
    private static JProgressBar progressBar = new JProgressBar(0, 100);
    private static JFrame frame;
    private static JComboBox<String> keyComboBox = new JComboBox<>();
    private static JTextField valueField = new JTextField();
    private static JTable configTable;
    private static JTable fileTable;
    private static JTextArea statsTextArea = new JTextArea();
    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";

    public static void main(String[] args) {
        // 创建并显示主窗口
        frame = createMainFrame();
        frame.setVisible(true);

        // 加载配置文件
        loadConfigFile(CONFIG_FILE);
    }

    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
             XWPFDocument doc = new XWPFDocument(fis)) {

            List<XWPFTable> tables = doc.getTables();
            for (XWPFTable table : tables) {
                for (XWPFTableRow row : table.getRows()) {
                    List<XWPFTableCell> cells = row.getTableCells();
                    if (cells.size() >= 5) {
                        String alias = cells.get(0).getText().trim();
                        List<String> keys = new ArrayList<>();
                        for (int i = 1; i < cells.size() - 1; i++) {
                            String key = cells.get(i).getText().trim();
                            if (!key.isEmpty() && !key.equals("-")) {
                                keys.add(key);
                            }
                        }
                        // Skip this row if no keys were found
                        if (keys.isEmpty()) {
                            continue;
                        }
                        String value = cells.get(cells.size() - 1).getText().trim();
                        aliasToKeysMap.put(alias, keys);
                        aliasToValueMap.put(alias, value);
                    }
                }
            }
        } catch (IOException e) {
            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
        }
    }


    private static void loadConfigFile(String configFile) {
        try (FileInputStream configFis = new FileInputStream(configFile);
             XWPFDocument configDoc = new XWPFDocument(configFis)) {

            List<XWPFTable> configTables = configDoc.getTables();

            for (XWPFTable table : configTables) {
                for (XWPFTableRow row : table.getRows()) {
                    if (row.getTableCells().size() >= 2) {
                        XWPFTableCell labelCell = row.getTableCells().get(0);
                        XWPFTableCell valueCell = row.getTableCells().get(1);
                        String key = labelCell.getText().trim();
                        String value = valueCell.getText().trim();
                        configMap.put(key, value);
                        configTableModel.addRow(new Object[]{key, value});
                        keyComboBox.addItem(key);
                    }
                }
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
        // After loading the original config file, load the equivalent keys file
        loadEquivalentKeysFile("Z:\\Desktop\\测试\\模板\\模板2.docx");

        // Merge the equivalent keys into the original config
        mergeEquivalentKeysToConfig();
    }
    private static void mergeEquivalentKeysToConfig() {
        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
            String alias = entry.getKey();
            String value = aliasToValueMap.get(alias);
            for (String key : entry.getValue()) {
                configMap.put(key, value);
                updateConfigTableModel(key, value);
            }
        }
    }
    private static void updateConfigTableModel(String key, String value) {
        boolean keyExists = false;
        for (int i = 0; i < configTableModel.getRowCount(); i++) {
            if (configTableModel.getValueAt(i, 0).equals(key)) {
                configTableModel.setValueAt(value, i, 1);
                keyExists = true;
                break;
            }
        }
        if (!keyExists) {
            configTableModel.addRow(new Object[]{key, value});
        }
    }


    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
        try (FileOutputStream fos = new FileOutputStream(configFile);
             XWPFDocument configDoc = new XWPFDocument()) {

            XWPFTable table = configDoc.createTable();
            for (Map.Entry<String, String> entry : configMap.entrySet()) {
                XWPFTableRow row = table.createRow();
                XWPFTableCell keyCell = row.addNewTableCell();
                keyCell.setText(entry.getKey());

                XWPFTableCell valueCell = row.addNewTableCell();
                valueCell.setText(entry.getValue());
            }
            configDoc.write(fos);
        } catch (IOException e) {
            System.err.println("An error occurred while saving the config file: " + e.getMessage());
        }
    }


    private static String convertDocToDocx(String sourceFile) {
        String docxFile = sourceFile.replace(".doc", ".docx");
        try {
            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
            doc.save(docxFile, SaveFormat.DOCX);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
        return docxFile;
    }

    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
             FileOutputStream fos = new FileOutputStream(outputFile);
             XWPFDocument document = new XWPFDocument(sourceFis)) {

            StringBuilder originalContent = new StringBuilder();
            StringBuilder modifiedContent = new StringBuilder();

            List<XWPFTable> tables = document.getTables();

            for (XWPFTable table : tables) {
                for (XWPFTableRow row : table.getRows()) {
                    for (int i = 0; i < row.getTableCells().size(); i++) {
                        XWPFTableCell cell = row.getTableCells().get(i);
                        String text = cell.getText().replaceAll("\\s+", "");
                        originalContent.append(text).append(" ");

                        if (configMap.containsKey(text)) {
                            String newValue = configMap.get(text);
                            modifiedContent.append(newValue).append(" ");
                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
                                nextCell.removeParagraph(0);
                                XWPFParagraph p = nextCell.addParagraph();
                                p.setAlignment(ParagraphAlignment.CENTER);
                                XWPFRun r = p.createRun();
                                r.setText(newValue);
                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
                            }
                        }
                    }
                }
            }

            for (XWPFParagraph paragraph : document.getParagraphs()) {
                List<XWPFRun> runs = paragraph.getRuns();
                for (int i = 0; i < runs.size(); i++) {
                    XWPFRun run = runs.get(i);
                    String text = run.getText(0);
                    if (text != null) {
                        originalContent.append(text).append(" ");
                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
                            String key = entry.getKey();
                            String value = entry.getValue();
                            if (text.trim().equals(key)) {
                                int j = i + 1;
                                while (j < runs.size()) {
                                    XWPFRun nextRun = runs.get(j);
                                    String nextText = nextRun.getText(0);
                                    if (nextText != null && !nextText.contains(":")) {
                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
                                        newRun.setText(value);
                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
                                        newRun.setFontFamily("仿宋_GB2312");
                                        newRun.setFontSize(14);
                                        paragraph.removeRun(j);
                                        break;
                                    }
                                    j++;
                                }
                                i = j;
                                break;
                            }
                        }
                    }
                }
            }

            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
            document.write(fos);
        }
    }

    private static void updateProgress(int processedFiles, int totalFiles) {
        int progress = (int) ((double) processedFiles / totalFiles * 100);
        progressBar.setValue(progress);
    }

    private static JFrame createMainFrame() {
        JFrame frame = new JFrame("文档处理工具");
        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
        frame.setSize(1200, 800);

        JPanel panel = new JPanel(new BorderLayout());
        frame.add(panel);

        JPanel configPanel = new JPanel(new BorderLayout());
        configPanel.setBorder(BorderFactory.createTitledBorder("配置文件映射"));
        configTable = new JTable(configTableModel);
        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);

        JPanel configInputPanel = new JPanel(new GridLayout(1, 4));
        configInputPanel.add(new JLabel("Key:"));
        configInputPanel.add(keyComboBox);
        configInputPanel.add(new JLabel("Value:"));
        configInputPanel.add(valueField);

        JButton addButton = new JButton("添加/更新");
        addButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                String key = (String) keyComboBox.getSelectedItem();
                String value = valueField.getText().trim();
                if (key != null && !key.isEmpty() && !value.isEmpty()) {
                    configMap.put(key, value);
                    boolean keyExists = false;
                    for (int i = 0; i < configTableModel.getRowCount(); i++) {
                        if (configTableModel.getValueAt(i, 0).equals(key)) {
                            configTableModel.setValueAt(value, i, 1);
                            keyExists = true;
                            break;
                        }
                    }
                    if (!keyExists) {
                        configTableModel.addRow(new Object[]{key, value});
                        keyComboBox.addItem(key);
                    }
                    saveConfigFile(CONFIG_FILE, configMap);
                }
            }
        });
        configInputPanel.add(addButton);

        configPanel.add(configInputPanel, BorderLayout.SOUTH);

        JPanel filePanel = new JPanel(new BorderLayout());
        filePanel.setBorder(BorderFactory.createTitledBorder("文件处理结果预览"));
        fileTable = new JTable(fileTableModel);
        fileTable.getSelectionModel().addListSelectionListener(new ListSelectionListener() {
            @Override
            public void valueChanged(ListSelectionEvent event) {
                if (!event.getValueIsAdjusting()) {
                    int selectedRow = fileTable.getSelectedRow();
                    if (selectedRow != -1) {
                        String fileName = (String) fileTableModel.getValueAt(selectedRow, 0);
                        String originalContent = (String) fileTableModel.getValueAt(selectedRow, 1);
                        String modifiedContent = (String) fileTableModel.getValueAt(selectedRow, 2);
                        displayFilePreview(fileName, originalContent, modifiedContent);
                    }
                }
            }
        });
        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);

        JButton refreshButton = new JButton("刷新预览");
        refreshButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                fileTableModel.setRowCount(0);
                processFiles();
            }
        });
        filePanel.add(refreshButton, BorderLayout.SOUTH);

        JPanel progressPanel = new JPanel(new BorderLayout());
        progressPanel.setBorder(BorderFactory.createTitledBorder("处理进度"));
        progressPanel.add(progressBar, BorderLayout.CENTER);

        JButton startButton = new JButton("开始执行");
        startButton.addActionListener(new ActionListener() {
            @Override
            public void actionPerformed(ActionEvent e) {
                processFiles();
                JOptionPane.showMessageDialog(frame, "处理完成！", "提示", JOptionPane.INFORMATION_MESSAGE);
            }
        });
        progressPanel.add(startButton, BorderLayout.SOUTH);

        JPanel statsPanel = new JPanel(new BorderLayout());
        statsPanel.setBorder(BorderFactory.createTitledBorder("统计分析"));
        statsTextArea.setEditable(false);
        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);

        panel.add(configPanel, BorderLayout.NORTH);
        panel.add(filePanel, BorderLayout.CENTER);
        panel.add(progressPanel, BorderLayout.SOUTH);
        panel.add(statsPanel, BorderLayout.EAST);

        return frame;
    }

    private static void displayFilePreview(String fileName, String originalContent, String modifiedContent) {
        JFrame previewFrame = new JFrame("文件预览 - " + fileName);
        previewFrame.setSize(600, 400);
        JTextArea previewTextArea = new JTextArea();
        previewTextArea.setEditable(false);
        previewTextArea.setText("原始内容:\n" + originalContent + "\n\n修改后内容:\n" + modifiedContent);
        previewFrame.add(new JScrollPane(previewTextArea));
        previewFrame.setVisible(true);
    }

    private static void processFiles() {
        File folder = new File("Z:\\Desktop\\测试\\in");
        File[] listOfFiles = folder.listFiles();
        if (listOfFiles == null) {
            return;
        }
        int totalFiles = listOfFiles.length;
        int processedFiles = 0;
        long startTime = System.currentTimeMillis();
        for (File file : listOfFiles) {
            if (file.isFile()) {
                String sourceFile = file.getAbsolutePath();
                String outputFile = "Z:\\Desktop\\测试\\out\\modified_" + file.getName();
                if (sourceFile.endsWith(".doc")) {
                    sourceFile = convertDocToDocx(sourceFile);
                }
                if (!sourceFile.endsWith(".docx")) {
                    continue;
                }
                try {
                    modifyDocument(sourceFile, outputFile);
                    processedFiles++;
                    updateProgress(processedFiles, totalFiles);
                } catch (IOException e) {
                    e.printStackTrace();
                }
            }
        }
        long endTime = System.currentTimeMillis();
        long duration = endTime - startTime;
        displayStats(totalFiles, processedFiles, duration);
    }

    private static void displayStats(int totalFiles, int processedFiles, long duration) {
        String stats = "总文件数: " + totalFiles + "\n已处理文件数: " + processedFiles + "\n处理时间: " + duration + " 毫秒\n";
        statsTextArea.setText(stats);
    }
}
//
//
//
//
//
//
//
//
//
//
//
//
//
//
//
//
//package org.example;
//
//import com.aspose.words.SaveFormat;
//import org.apache.poi.xwpf.usermodel.*;
//
//import javax.swing.*;
//import javax.swing.table.DefaultTableModel;
//import java.awt.*;
//import java.awt.event.ActionEvent;
//import java.awt.event.ActionListener;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.*;
//import java.util.List;
//
//public class WordModifier {
//    private static Map<String, String> configMap = new LinkedHashMap<>();
//    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
//    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();
//    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//    private static JProgressBar progressBar = new JProgressBar(0, 100);
//    private static JFrame frame;
//    private static JComboBox<String> keyComboBox = new JComboBox<>();
//    private static JTextField valueField = new JTextField();
//    private static JTable configTable;
//    private static JTable fileTable;
//    private static JTextArea statsTextArea = new JTextArea();
//    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//    private static final String EQUIVALENT_KEYS_FILE = "Z:\\Desktop\\测试\\模板\\模板2.docx";
//
//    public static void main(String[] args) {
//        // 创建并显示主窗口
//        frame = createMainFrame();
//        frame.setVisible(true);
//
//        // 加载配置文件
//        loadConfigFile(CONFIG_FILE);
//
//        // 加载等价键配置文件
//        loadEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
//
//        // 合并等价键到原有配置文件中
//        mergeEquivalentKeysToConfig();
//
//        // 显示等价键设置界面
//        displayEquivalentKeysUI();
//    }
//
//    private static void loadConfigFile(String configFile) {
//        try (FileInputStream configFis = new FileInputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//
//            List<XWPFTable> configTables = configDoc.getTables();
//
//            for (XWPFTable table : configTables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    if (row.getTableCells().size() >= 2) {
//                        XWPFTableCell labelCell = row.getTableCells().get(0);
//                        XWPFTableCell valueCell = row.getTableCells().get(1);
//                        String key = labelCell.getText().trim();
//                        String value = valueCell.getText().trim();
//                        configMap.put(key, value);
//                        configTableModel.addRow(new Object[]{key, value});
//                        keyComboBox.addItem(key);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
//        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
//             XWPFDocument doc = new XWPFDocument(fis)) {
//
//            List<XWPFTable> tables = doc.getTables();
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    List<XWPFTableCell> cells = row.getTableCells();
//                    if (cells.size() >= 5) {
//                        String alias = cells.get(0).getText().trim();
//                        List<String> keys = new ArrayList<>();
//                        for (int i = 1; i < cells.size() - 1; i++) {
//                            String key = cells.get(i).getText().trim();
//                            if (!key.isEmpty() && !key.equals("-")) {
//                                keys.add(key);
//                            }
//                        }
//                        String value = cells.get(cells.size() - 1).getText().trim();
//                        aliasToKeysMap.put(alias, keys);
//                        aliasToValueMap.put(alias, value);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
//        }
//    }
//
//    private static void mergeEquivalentKeysToConfig() {
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                configMap.put(key, value);
//                updateConfigTableModel(key, value);
//            }
//        }
//    }
//
//    private static void updateConfigTableModel(String key, String value) {
//        boolean keyExists = false;
//        for (int i = 0; i < configTableModel.getRowCount(); i++) {
//            if (configTableModel.getValueAt(i, 0).equals(key)) {
//                configTableModel.setValueAt(value, i, 1);
//                keyExists = true;
//                break;
//            }
//        }
//        if (!keyExists) {
//            configTableModel.addRow(new Object[]{key, value});
//            keyComboBox.addItem(key);
//        }
//    }
//
//    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//        try (FileOutputStream fos = new FileOutputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument()) {
//
//            XWPFTable table = configDoc.createTable();
//            boolean firstRow = true;
//            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                XWPFTableRow row;
//                if (firstRow) {
//                    row = table.getRow(0); // 使用已经存在的第一行
//                    firstRow = false;
//                } else {
//                    row = table.createRow();
//                }
//                row.getCell(0).setText(entry.getKey());
//                row.addNewTableCell().setText(entry.getValue()); // addNewTableCell to create a new cell
//            }
//            configDoc.write(fos);
//        } catch (IOException e) {
//            // 这里可以使用更合适的异常处理方式
//            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//        }
//    }
//
//    private static String convertDocToDocx(String sourceFile) {
//        String docxFile = sourceFile.replace(".doc", ".docx");
//        try {
//            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//            doc.save(docxFile, SaveFormat.DOCX);
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
//        return docxFile;
//    }
//
//    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//             FileOutputStream fos = new FileOutputStream(outputFile);
//             XWPFDocument document = new XWPFDocument(sourceFis)) {
//
//            StringBuilder originalContent = new StringBuilder();
//            StringBuilder modifiedContent = new StringBuilder();
//
//            List<XWPFTable> tables = document.getTables();
//
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    for (int i = 0; i < row.getTableCells().size(); i++) {
//                        XWPFTableCell cell = row.getTableCells().get(i);
//                        String text = cell.getText().replaceAll("\\s+", "");
//                        originalContent.append(text).append(" ");
//
//                        if (configMap.containsKey(text)) {
//                            String newValue = configMap.get(text);
//                            modifiedContent.append(newValue).append(" ");
//                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//                                nextCell.removeParagraph(0);
//                                XWPFParagraph p = nextCell.addParagraph();
//                                p.setAlignment(ParagraphAlignment.CENTER);
//                                XWPFRun r = p.createRun();
//                                r.setText(newValue);
//                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                            }
//                        }
//                    }
//                }
//            }
//
//            for (XWPFParagraph paragraph : document.getParagraphs()) {
//                List<XWPFRun> runs = paragraph.getRuns();
//                for (int i = 0; i < runs.size(); i++) {
//                    XWPFRun run = runs.get(i);
//                    String text = run.getText(0);
//                    if (text != null) {
//                        originalContent.append(text).append(" ");
//                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                            String key = entry.getKey();
//                            String value = entry.getValue();
//                            if (text.trim().equals(key)) {
//                                int j = i + 1;
//                                while (j < runs.size()) {
//                                    XWPFRun nextRun = runs.get(j);
//                                    String nextText = nextRun.getText(0);
//                                    if (nextText != null && !nextText.contains(":")) {
//                                        XWPFRun newRun = paragraph.insertNewRun(j + 1);
//                                        newRun.setText(value);
//                                        newRun.setUnderline(UnderlinePatterns.SINGLE);
//                                        newRun.setFontFamily("仿宋");
//                                        newRun.setFontSize(12);
//                                        break;
//                                    }
//                                    j++;
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//
//            document.write(fos);
//        }
//    }
//
//    private static void displayEquivalentKeysUI() {
//        JFrame aliasFrame = new JFrame("设置等价键值");
//        aliasFrame.setSize(400, 300);
//        JPanel aliasPanel = new JPanel(new BorderLayout());
//
//        DefaultTableModel aliasTableModel = new DefaultTableModel(new Object[]{"别名", "键", "值"}, 0);
//        JTable aliasTable = new JTable(aliasTableModel);
//
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                aliasTableModel.addRow(new Object[]{alias, key, value});
//            }
//        }
//
//        aliasPanel.add(new JScrollPane(aliasTable), BorderLayout.CENTER);
//
//        JButton saveButton = new JButton("保存");
//        saveButton.addActionListener(e -> {
//            // 处理保存逻辑
//            // 你可以添加代码来保存用户在界面上修改的值
//        });
//        aliasPanel.add(saveButton, BorderLayout.SOUTH);
//
//        aliasFrame.add(aliasPanel);
//        aliasFrame.setVisible(true);
//    }
//
//    private static JFrame createMainFrame() {
//        JFrame frame = new JFrame("Word Modifier");
//        frame.setSize(800, 600);
//        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        frame.setLayout(new BorderLayout());
//
//        JTabbedPane tabbedPane = new JTabbedPane();
//
//        // Configuration Panel
//        JPanel configPanel = new JPanel(new BorderLayout());
//        configTable = new JTable(configTableModel);
//        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Configuration", configPanel);
//
//        // File Panel
//        JPanel filePanel = new JPanel(new BorderLayout());
//        fileTable = new JTable(fileTableModel);
//        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Files", filePanel);
//
//        frame.add(tabbedPane, BorderLayout.CENTER);
//
//        // Control Panel
//        JPanel controlPanel = new JPanel();
//        controlPanel.add(new JLabel("Key:"));
//        controlPanel.add(keyComboBox);
//        controlPanel.add(new JLabel("Value:"));
//        valueField.setPreferredSize(new Dimension(100, 24));
//        controlPanel.add(valueField);
//        JButton updateButton = new JButton("Update");
//        updateButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                String key = (String) keyComboBox.getSelectedItem();
//                String value = valueField.getText();
//                if (key != null && !value.isEmpty()) {
//                    configMap.put(key, value);
//                    updateConfigTableModel(key, value);
//                    saveConfigFile(CONFIG_FILE, configMap);
//                }
//            }
//        });
//        controlPanel.add(updateButton);
//
//        JButton processFilesButton = new JButton("Process Files");
//        processFilesButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                processFiles();
//            }
//        });
//        controlPanel.add(processFilesButton);
//        frame.add(controlPanel, BorderLayout.SOUTH);
//
//        frame.add(progressBar, BorderLayout.NORTH);
//
//        // Statistics Panel
//        JPanel statsPanel = new JPanel(new BorderLayout());
//        statsTextArea.setEditable(false);
//        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//        tabbedPane.addTab("Statistics", statsPanel);
//
//        return frame;
//    }
//
//    private static void processFiles() {
//        JFileChooser fileChooser = new JFileChooser();
//        fileChooser.setMultiSelectionEnabled(true);
//        int returnValue = fileChooser.showOpenDialog(frame);
//        if (returnValue == JFileChooser.APPROVE_OPTION) {
//            File[] selectedFiles = fileChooser.getSelectedFiles();
//            progressBar.setMaximum(selectedFiles.length);
//            progressBar.setValue(0);
//            int processedFiles = 0;
//            for (File file : selectedFiles) {
//                String sourceFile = file.getAbsolutePath();
//                String outputFile = sourceFile.replace(".doc", "_modified.docx").replace(".docx", "_modified.docx");
//
//                // Convert .doc files to .docx
//                if (sourceFile.toLowerCase().endsWith(".doc")) {
//                    sourceFile = convertDocToDocx(sourceFile);
//                }
//
//                try {
//                    modifyDocument(sourceFile, outputFile);
//                    processedFiles++;
//                    progressBar.setValue(processedFiles);
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
//            showStats();
//        }
//    }
//
//    private static void showStats() {
//        StringBuilder stats = new StringBuilder();
//        stats.append("处理的文件数: ").append(fileTableModel.getRowCount()).append("\n");
//        for (int i = 0; i < fileTableModel.getRowCount(); i++) {
//            String fileName = (String) fileTableModel.getValueAt(i, 0);
//            String originalContent = (String) fileTableModel.getValueAt(i, 1);
//            String modifiedContent = (String) fileTableModel.getValueAt(i, 2);
//            stats.append("文件: ").append(fileName).append("\n")
//                    .append("原始内容: ").append(originalContent).append("\n")
//                    .append("修改后内容: ").append(modifiedContent).append("\n");
//        }
//        statsTextArea.setText(stats.toString());
//    }
//}






//
//package org.example;
//
//import com.aspose.words.SaveFormat;
//import org.apache.poi.xwpf.usermodel.*;
//
//import javax.swing.*;
//import javax.swing.table.DefaultTableModel;
//import java.awt.*;
//import java.awt.event.ActionEvent;
//import java.awt.event.ActionListener;
//import java.io.File;
//import java.io.FileInputStream;
//import java.io.FileOutputStream;
//import java.io.IOException;
//import java.util.*;
//import java.util.List;
//
//public class WordModifier {
//    private static Map<String, String> configMap = new LinkedHashMap<>();
//    private static Map<String, List<String>> aliasToKeysMap = new LinkedHashMap<>();
//    private static Map<String, String> aliasToValueMap = new LinkedHashMap<>();
//    private static DefaultTableModel configTableModel = new DefaultTableModel(new Object[]{"Key", "Value"}, 0);
//    private static DefaultTableModel fileTableModel = new DefaultTableModel(new Object[]{"File", "Original Content", "Modified Content"}, 0);
//    private static JProgressBar progressBar = new JProgressBar(0, 100);
//    private static JFrame frame;
//    private static JComboBox<String> keyComboBox = new JComboBox<>();
//    private static JTextField valueField = new JTextField();
//    private static JTable configTable;
//    private static JTable fileTable;
//    private static JTextArea statsTextArea = new JTextArea();
//    private static final String CONFIG_FILE = "Z:\\Desktop\\测试\\模板\\模板.docx";
//    private static final String EQUIVALENT_KEYS_FILE = "Z:\\Desktop\\测试\\模板\\模板2.docx";
//
//    public static void main(String[] args) {
//        // 创建并显示主窗口
//        frame = createMainFrame();
//        frame.setVisible(true);
//
//        // 加载配置文件
//        loadConfigFile(CONFIG_FILE);
//
//        // 加载等价键配置文件
//        loadEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
//
//        // 合并等价键到原有配置文件中
//        mergeEquivalentKeysToConfig();
//
//        // 显示等价键设置界面
//        displayEquivalentKeysUI();
//    }
//
//    private static void loadConfigFile(String configFile) {
//        try (FileInputStream configFis = new FileInputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument(configFis)) {
//
//            List<XWPFTable> configTables = configDoc.getTables();
//
//            for (XWPFTable table : configTables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    if (row.getTableCells().size() >= 2) {
//                        XWPFTableCell labelCell = row.getTableCells().get(0);
//                        XWPFTableCell valueCell = row.getTableCells().get(1);
//                        String key = labelCell.getText().trim();
//                        String value = valueCell.getText().trim();
//                        configMap.put(key, value);
//                        configTableModel.addRow(new Object[]{key, value});
//                        keyComboBox.addItem(key);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            e.printStackTrace();
//        }
//    }
//
//    private static void loadEquivalentKeysFile(String equivalentKeysFile) {
//        try (FileInputStream fis = new FileInputStream(equivalentKeysFile);
//             XWPFDocument doc = new XWPFDocument(fis)) {
//
//            List<XWPFTable> tables = doc.getTables();
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    List<XWPFTableCell> cells = row.getTableCells();
//                    if (cells.size() >= 5) {
//                        String alias = cells.get(0).getText().trim();
//                        List<String> keys = new ArrayList<>();
//                        for (int i = 1; i < cells.size() - 1; i++) {
//                            String key = cells.get(i).getText().trim();
//                            if (!key.isEmpty() && !key.equals("-")) {
//                                keys.add(key);
//                            }
//                        }
//                        String value = cells.get(cells.size() - 1).getText().trim();
//                        aliasToKeysMap.put(alias, keys);
//                        aliasToValueMap.put(alias, value);
//                    }
//                }
//            }
//        } catch (IOException e) {
//            System.err.println("An error occurred while loading the equivalent keys file: " + e.getMessage());
//        }
//    }
//
//    private static void mergeEquivalentKeysToConfig() {
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                configMap.put(key, value);
//                updateConfigTableModel(key, value);
//            }
//        }
//    }
//
//    private static void updateConfigTableModel(String key, String value) {
//        boolean keyExists = false;
//        for (int i = 0; i < configTableModel.getRowCount(); i++) {
//            if (configTableModel.getValueAt(i, 0).equals(key)) {
//                configTableModel.setValueAt(value, i, 1);
//                keyExists = true;
//                break;
//            }
//        }
//        if (!keyExists) {
//            configTableModel.addRow(new Object[]{key, value});
//            keyComboBox.addItem(key);
//        }
//    }
//
//    private static void saveConfigFile(String configFile, Map<String, String> configMap) {
//        try (FileOutputStream fos = new FileOutputStream(configFile);
//             XWPFDocument configDoc = new XWPFDocument()) {
//
//            XWPFTable table = configDoc.createTable();
//            boolean firstRow = true;
//            for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                XWPFTableRow row;
//                if (firstRow) {
//                    row = table.getRow(0); // 使用已经存在的第一行
//                    firstRow = false;
//                } else {
//                    row = table.createRow();
//                }
//                row.getCell(0).setText(entry.getKey());
//                row.addNewTableCell().setText(entry.getValue()); // addNewTableCell to create a new cell
//            }
//            configDoc.write(fos);
//        } catch (IOException e) {
//            // 这里可以使用更合适的异常处理方式
//            System.err.println("An error occurred while saving the config file: " + e.getMessage());
//        }
//    }
//
//    private static String convertDocToDocx(String sourceFile) {
//        String docxFile = sourceFile.replace(".doc", ".docx");
//        try {
//            com.aspose.words.Document doc = new com.aspose.words.Document(sourceFile);
//            doc.save(docxFile, SaveFormat.DOCX);
//        } catch (Exception e) {
//            throw new RuntimeException(e);
//        }
//        return docxFile;
//    }
//
//    private static void modifyDocument(String sourceFile, String outputFile) throws IOException {
//        try (FileInputStream sourceFis = new FileInputStream(sourceFile);
//             FileOutputStream fos = new FileOutputStream(outputFile);
//             XWPFDocument document = new XWPFDocument(sourceFis)) {
//
//            StringBuilder originalContent = new StringBuilder();
//            StringBuilder modifiedContent = new StringBuilder();
//
//            List<XWPFTable> tables = document.getTables();
//
//            for (XWPFTable table : tables) {
//                for (XWPFTableRow row : table.getRows()) {
//                    for (int i = 0; i < row.getTableCells().size(); i++) {
//                        XWPFTableCell cell = row.getTableCells().get(i);
//                        String text = cell.getText().replaceAll("\\s+", "");
//                        originalContent.append(text).append(" ");
//
//                        if (configMap.containsKey(text)) {
//                            String newValue = configMap.get(text);
//                            modifiedContent.append(newValue).append(" ");
//                            if (i + 1 < row.getTableCells().size() && newValue != null && !newValue.isEmpty()) {
//                                XWPFTableCell nextCell = row.getTableCells().get(i + 1);
//                                nextCell.removeParagraph(0);
//                                XWPFParagraph p = nextCell.addParagraph();
//                                p.setAlignment(ParagraphAlignment.CENTER);
//                                XWPFRun r = p.createRun();
//                                r.setText(newValue);
//                                nextCell.setVerticalAlignment(XWPFTableCell.XWPFVertAlign.CENTER);
//                            }
//                        }
//                    }
//                }
//            }
//
//            for (XWPFParagraph paragraph : document.getParagraphs()) {
//                List<XWPFRun> runs = paragraph.getRuns();
//                for (int i = 0; i < runs.size(); i++) {
//                    XWPFRun run = runs.get(i);
//                    String text = run.getText(0);
//                    if (text != null) {
//                        originalContent.append(text).append(" ");
//                        for (Map.Entry<String, String> entry : configMap.entrySet()) {
//                            String key = entry.getKey();
//                            String value = entry.getValue();
//                            if (text.trim().equals(key)) {
//                                int j = i + 1;
//                                while (j < runs.size()) {
//                                    XWPFRun nextRun = runs.get(j);
//                                    String nextText = nextRun.getText(0);
//                                    if (nextText != null && nextText.trim().equals(value)) {
//                                        run.setText(value, 0);
//                                        modifiedContent.append(value).append(" ");
//                                        nextRun.setText("", 0);
//                                        break;
//                                    }
//                                    j++;
//                                }
//                            }
//                        }
//                    }
//                }
//            }
//
//            fileTableModel.addRow(new Object[]{sourceFile, originalContent.toString(), modifiedContent.toString()});
//
//            document.write(fos);
//        }
//    }
//
//    private static void displayEquivalentKeysUI() {
//        JFrame aliasFrame = new JFrame("设置等价键值");
//        aliasFrame.setSize(400, 300);
//        JPanel aliasPanel = new JPanel(new BorderLayout());
//
//        DefaultTableModel aliasTableModel = new DefaultTableModel(new Object[]{"别名", "键", "值"}, 0);
//        JTable aliasTable = new JTable(aliasTableModel);
//
//        for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//            String alias = entry.getKey();
//            String value = aliasToValueMap.get(alias);
//            for (String key : entry.getValue()) {
//                aliasTableModel.addRow(new Object[]{alias, key, value});
//            }
//        }
//
//        aliasPanel.add(new JScrollPane(aliasTable), BorderLayout.CENTER);
//
//        JButton saveButton = new JButton("保存");
//        saveButton.addActionListener(e -> {
//            // 处理保存逻辑
//            aliasToKeysMap.clear();
//            aliasToValueMap.clear();
//            for (int i = 0; i < aliasTableModel.getRowCount(); i++) {
//                String alias = (String) aliasTableModel.getValueAt(i, 0);
//                String key = (String) aliasTableModel.getValueAt(i, 1);
//                String value = (String) aliasTableModel.getValueAt(i, 2);
//
//                aliasToKeysMap.computeIfAbsent(alias, k -> new ArrayList<>()).add(key);
//                aliasToValueMap.put(alias, value);
//            }
//
//            saveEquivalentKeysFile(EQUIVALENT_KEYS_FILE);
//            mergeEquivalentKeysToConfig();
//        });
//        aliasPanel.add(saveButton, BorderLayout.SOUTH);
//
//        aliasFrame.add(aliasPanel);
//        aliasFrame.setVisible(true);
//    }
//
//    private static void saveEquivalentKeysFile(String equivalentKeysFile) {
//        try (FileOutputStream fos = new FileOutputStream(equivalentKeysFile);
//             XWPFDocument doc = new XWPFDocument()) {
//
//            XWPFTable table = doc.createTable();
//            for (Map.Entry<String, List<String>> entry : aliasToKeysMap.entrySet()) {
//                XWPFTableRow row = table.createRow();
//                row.getCell(0).setText(entry.getKey());
//                int cellIndex = 1;
//                for (String key : entry.getValue()) {
//                    if (cellIndex >= row.getTableCells().size()) {
//                        row.addNewTableCell().setText(key);
//                    } else {
//                        row.getCell(cellIndex).setText(key);
//                    }
//                    cellIndex++;
//                }
//                if (cellIndex >= row.getTableCells().size()) {
//                    row.addNewTableCell().setText(aliasToValueMap.get(entry.getKey()));
//                } else {
//                    row.getCell(cellIndex).setText(aliasToValueMap.get(entry.getKey()));
//                }
//            }
//
//            doc.write(fos);
//        } catch (IOException e) {
//            System.err.println("An error occurred while saving the equivalent keys file: " + e.getMessage());
//        }
//    }
//
//    private static JFrame createMainFrame() {
//        JFrame frame = new JFrame("Word Modifier");
//        frame.setSize(800, 600);
//        frame.setDefaultCloseOperation(JFrame.EXIT_ON_CLOSE);
//        frame.setLayout(new BorderLayout());
//
//        JTabbedPane tabbedPane = new JTabbedPane();
//
//        // Configuration Panel
//        JPanel configPanel = new JPanel(new BorderLayout());
//        configTable = new JTable(configTableModel);
//        configPanel.add(new JScrollPane(configTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Configuration", configPanel);
//
//        // File Panel
//        JPanel filePanel = new JPanel(new BorderLayout());
//        fileTable = new JTable(fileTableModel);
//        filePanel.add(new JScrollPane(fileTable), BorderLayout.CENTER);
//        tabbedPane.addTab("Files", filePanel);
//
//        frame.add(tabbedPane, BorderLayout.CENTER);
//
//        // Control Panel
//        JPanel controlPanel = new JPanel();
//        controlPanel.add(new JLabel("Key:"));
//        controlPanel.add(keyComboBox);
//        controlPanel.add(new JLabel("Value:"));
//        valueField.setPreferredSize(new Dimension(100, 24));
//        controlPanel.add(valueField);
//        JButton updateButton = new JButton("Update");
//        updateButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                String key = (String) keyComboBox.getSelectedItem();
//                String value = valueField.getText();
//                if (key != null && !value.isEmpty()) {
//                    configMap.put(key, value);
//                    updateConfigTableModel(key, value);
//                    saveConfigFile(CONFIG_FILE, configMap);
//                }
//            }
//        });
//        controlPanel.add(updateButton);
//
//        JButton processFilesButton = new JButton("Process Files");
//        processFilesButton.addActionListener(new ActionListener() {
//            @Override
//            public void actionPerformed(ActionEvent e) {
//                processFiles();
//            }
//        });
//        controlPanel.add(processFilesButton);
//        frame.add(controlPanel, BorderLayout.SOUTH);
//
//        frame.add(progressBar, BorderLayout.NORTH);
//
//        // Statistics Panel
//        JPanel statsPanel = new JPanel(new BorderLayout());
//        statsTextArea.setEditable(false);
//        statsPanel.add(new JScrollPane(statsTextArea), BorderLayout.CENTER);
//        tabbedPane.addTab("Statistics", statsPanel);
//
//        return frame;
//    }
//
//    private static void processFiles() {
//        JFileChooser fileChooser = new JFileChooser();
//        fileChooser.setMultiSelectionEnabled(true);
//        int returnValue = fileChooser.showOpenDialog(frame);
//        if (returnValue == JFileChooser.APPROVE_OPTION) {
//            File[] selectedFiles = fileChooser.getSelectedFiles();
//            progressBar.setMaximum(selectedFiles.length);
//            progressBar.setValue(0);
//            int processedFiles = 0;
//            for (File file : selectedFiles) {
//                String sourceFile = file.getAbsolutePath();
//                String outputFile = sourceFile.replace(".doc", "_modified.docx").replace(".docx", "_modified.docx");
//
//                // Convert .doc files to .docx
//                if (sourceFile.toLowerCase().endsWith(".doc")) {
//                    sourceFile = convertDocToDocx(sourceFile);
//                }
//
//                try {
//                    modifyDocument(sourceFile, outputFile);
//                    processedFiles++;
//                    progressBar.setValue(processedFiles);
//                } catch (IOException e) {
//                    e.printStackTrace();
//                }
//            }
//            showStats();
//        }
//    }
//
//    private static void showStats() {
//        StringBuilder stats = new StringBuilder();
//        stats.append("处理的文件数: ").append(fileTableModel.getRowCount()).append("\n");
//        for (int i = 0; i < fileTableModel.getRowCount(); i++) {
//            String fileName = (String) fileTableModel.getValueAt(i, 0);
//            String originalContent = (String) fileTableModel.getValueAt(i, 1);
//            String modifiedContent = (String) fileTableModel.getValueAt(i, 2);
//            stats.append("文件: ").append(fileName).append("\n")
//                    .append("原始内容: ").append(originalContent).append("\n")
//                    .append("修改后内容: ").append(modifiedContent).append("\n");
//        }
//        statsTextArea.setText(stats.toString());
//    }
//}
//
