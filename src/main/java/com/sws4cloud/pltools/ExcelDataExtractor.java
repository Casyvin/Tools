package com.sws4cloud.pltools;

import org.apache.poi.ss.usermodel.*;

import java.io.*;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;


public class ExcelDataExtractor {

    // 添加静态日志回调变量
    private static LogCallback currentLogCallback = null;

    // 源文件列索引
    private static final int SRC_COL_ID_PALLET = 16;         // Q - ID PALLET
    private static final int SRC_COL_LABEL = 3;              // D - LABEL
    private static final int SRC_COL_VARIETY = 2;            // C - VARIETY
    private static final int SRC_COL_SIZE = 5;               // F - SIZE
    private static final int SRC_COL_NET_WEIGHT = 7;         // H - NET WEIGHT
    private static final int SRC_COL_QUANTITY = 6;           // G - QUANTITY OF TRAYS
    private static final int SRC_COL_CSG = 10;               // K - CSG
    private static final int SRC_COL_CSP = 9;                // J - CSP
    private static final int SRC_COL_PACKING_DATE = 15;      // P - PACKING DATE
    private static final int SRC_COL_CAT = 4;                // E - CAT
    private static final int SRC_COL_TEMP_RECORDER = 18;     // S - TEMPERATURE RECORDER

    // 目标文件列索引
    private static final int TGT_COL_SPECIES = 0;            // A - Species
    private static final int TGT_COL_PALLET_NO = 1;          // B - Pallet No.
    private static final int TGT_COL_BRAND = 2;              // C - Brand
    private static final int TGT_COL_VARIETY = 3;            // D - Variety
    private static final int TGT_COL_SIZE = 4;               // E - Size
    private static final int TGT_COL_NW = 5;                 // F - N.W
    private static final int TGT_COL_CASES = 6;              // G - CASES
    private static final int TGT_COL_TOTAL_NW = 7;           // H - TOTAL N.W
    private static final int TGT_COL_CSG_CODE = 8;           // I - CSG Code
    private static final int TGT_COL_CSP_CODE = 9;           // J - CSP Code
    private static final int TGT_COL_PACKING_DATE = 10;      // K - Packing Date
    private static final int TGT_COL_CATEGORY = 11;          // L - Category
    private static final int TGT_COL_THERMOGRAPH = 12;       // M - Thermograph
    private static final int TGT_COL_TOTAL_CASES_PALLET = 13; // N - Total cases per pallet
    private static final int TGT_COL_PALLETS = 14;           // O - Pallets

    // 目标文件数据起始行（从第15行开始，索引14）
    private static final int TARGET_START_ROW = 14;

    // 在 ExcelDataExtractor 类中添加以下内容：

    /**
     * 日志输出回调接口
     */
    public interface LogCallback {
        void logMessage(String message);

        void logError(String message);
    }

    /**
     * 执行数据提取和转换的主要方法
     *
     * @param templateFilePath 模板文件路径
     * @param sourceDirPath    源文件目录路径
     * @param outputDirPath    输出目录路径
     * @param logCallback      日志回调接口
     */
    public static void executeDataExtraction(String templateFilePath, String sourceDirPath,
                                             String outputDirPath, LogCallback logCallback) {
        try {
            logCallback.logMessage("=== Excel数据迁移工具 ===");
            logCallback.logMessage("模板文件: " + templateFilePath);
            logCallback.logMessage("源文件目录: " + sourceDirPath);
            logCallback.logMessage("输出目录: " + outputDirPath);

            // 在创建输出目录时确保路径格式正确
            File outputDir = new File(outputDirPath);
            if (!outputDir.exists()) {
                outputDir.mkdirs();
            }
            // 确保路径以分隔符结尾
            if (!outputDirPath.endsWith(File.separator)) {
                outputDirPath = outputDirPath + File.separator;
            }

            // 获取源目录中的所有Excel文件
            File sourceDir = new File(sourceDirPath);
            if (!sourceDir.exists() || !sourceDir.isDirectory()) {
                logCallback.logError("源目录不存在或不是目录: " + sourceDirPath);
                return;
            }

            File[] sourceFiles = sourceDir.listFiles((dir, name) ->
                    name.toLowerCase().endsWith(".xlsx") && !name.startsWith("~$"));
            if (sourceFiles == null || sourceFiles.length == 0) {
                logCallback.logMessage("源目录中没有找到Excel文件");
                return;
            }
            logCallback.logMessage("找到 " + sourceFiles.length + " 个Excel文件");

            // 循环处理每个文件
            for (int i = 0; i < sourceFiles.length; i++) {
                File sourceFile = sourceFiles[i];
                logCallback.logMessage("\n[" + (i + 1) + "/" + sourceFiles.length + "] 处理文件: " + sourceFile.getName());

                try {
                    String sourceFilePath = sourceFile.getAbsolutePath();
                    String outputFilePath = outputDirPath + sourceFile.getName();

                    // 1. 从源文件提取数据
                    logCallback.logMessage("  1. 从源文件提取数据...");
                    List<DataRow> sourceData = extractDataFromSource(sourceFilePath);
                    logCallback.logMessage("     提取到 " + sourceData.size() + " 行数据");

                    // 2. 计算每个托盘的汇总信息
                    logCallback.logMessage("  2. 计算托盘汇总信息...");
                    Map<String, Integer> palletTotals = calculatePalletTotals(sourceData);

                    // 3. 将数据写入模板
                    logCallback.logMessage("  3. 将数据写入模板文件...");

                    // 修改 writeDataToTemplate 调用
                    writeDataToTemplate(sourceData, palletTotals, templateFilePath, outputFilePath, logCallback);

                    logCallback.logMessage("  处理完成！输出文件: " + outputFilePath);

                } catch (Exception e) {
                    logCallback.logError("  处理文件 " + sourceFile.getName() + " 时发生错误: " + e.getMessage());
                    e.printStackTrace();
                }
            }

            logCallback.logMessage("\n所有文件处理完成！");

        } catch (Exception e) {
            logCallback.logError("处理过程中发生错误: " + e.getMessage());
            e.printStackTrace();
        }
    }

    /**
     * 从源文件提取数据（跳过标题行）
     */
    public static List<DataRow> extractDataFromSource(String sourceFilePath) throws IOException {
        List<DataRow> dataList = new ArrayList<>();

        try (FileInputStream fis = new FileInputStream(sourceFilePath);
             Workbook workbook = WorkbookFactory.create(fis)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 从第14行开始（索引13），但这是标题行，实际数据从第15行开始（索引14）
            int startRow = 14; // Excel第15行（索引14）

            for (int rowNum = startRow; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row row = sheet.getRow(rowNum);
                if (row == null) continue;

                // 跳过空行
                if (isRowEmpty(row)) continue;

                // 提取数据
                DataRow dataRow = extractRowData(row);
                if (dataRow != null) {
                    dataList.add(dataRow);
                }
            }
        }

        return dataList;
    }

    /**
     * 从单行提取数据
     */
    private static DataRow extractRowData(Row row) {
        try {
            // 获取各个字段的值
            String idPallet = getCellValue(row.getCell(SRC_COL_ID_PALLET));

            // 跳过没有托盘号的行
            if (idPallet == null || idPallet.trim().isEmpty()) {
                return null;
            }

            DataRow dataRow = new DataRow();
            dataRow.idPallet = idPallet;
            dataRow.label = getCellValue(row.getCell(SRC_COL_LABEL));
            dataRow.variety = getCellValue(row.getCell(SRC_COL_VARIETY));
            dataRow.size = getCellValue(row.getCell(SRC_COL_SIZE));
            dataRow.netWeight = getCellValue(row.getCell(SRC_COL_NET_WEIGHT));
            dataRow.quantity = getCellValue(row.getCell(SRC_COL_QUANTITY));
            dataRow.csg = getCellValue(row.getCell(SRC_COL_CSG));
            dataRow.csp = getCellValue(row.getCell(SRC_COL_CSP));
            dataRow.packingDate = getCellValue(row.getCell(SRC_COL_PACKING_DATE));
            dataRow.cat = getCellValue(row.getCell(SRC_COL_CAT));
            dataRow.tempRecorder = getCellValue(row.getCell(SRC_COL_TEMP_RECORDER));

            return dataRow;

        } catch (Exception e) {
            //System.err.println("提取行数据时出错: " + e.getMessage());
            // 替换 System.err.println 为日志回调
            if (currentLogCallback != null) {
                currentLogCallback.logError("提取行数据时出错: " + e.getMessage());
            }
            return null;
        }
    }

    /**
     * 计算每个托盘的汇总信息
     */
    private static Map<String, Integer> calculatePalletTotals(List<DataRow> dataList) {
        Map<String, Integer> palletTotals = new HashMap<>();

        for (DataRow row : dataList) {
            try {
                int quantity = parseQuantity(row.quantity);
                String palletNo = row.idPallet;

                palletTotals.put(palletNo, palletTotals.getOrDefault(palletNo, 0) + quantity);
            } catch (Exception e) {
                //System.err.println("计算托盘汇总时出错: " + e.getMessage());
                // 替换 System.err.println 为日志回调
                if (currentLogCallback != null) {
                    currentLogCallback.logError("计算托盘汇总时出错: " + e.getMessage());
                }
            }
        }

        return palletTotals;
    }

    /**
     * 解析数量
     */
    private static int parseQuantity(String quantityStr) {
        if (quantityStr == null || quantityStr.trim().isEmpty()) {
            return 0;
        }

        try {
            // 移除逗号和空格
            quantityStr = quantityStr.replace(",", "").replace(" ", "");
            return Integer.parseInt(quantityStr);
        } catch (NumberFormatException e) {
            //System.err.println("解析数量失败: " + quantityStr);
            // 替换 System.err.println 为日志回调
            if (currentLogCallback != null) {
                currentLogCallback.logError("解析数量失败: " + quantityStr);
            }
            return 0;
        }
    }

    /**
     * 解析净重
     */
    private static double parseNetWeight(String weightStr) {
        if (weightStr == null || weightStr.trim().isEmpty()) {
            return 0.0;
        }

        try {
            // 替换逗号为点，并解析
            weightStr = weightStr.replace(',', '.').replace(" ", "");
            return Double.parseDouble(weightStr);
        } catch (NumberFormatException e) {
            //System.err.println("解析净重失败: " + weightStr);
            // 替换 System.err.println 为日志回调
            if (currentLogCallback != null) {
                currentLogCallback.logError("解析净重失败: " + weightStr);
            }
            return 0.0;
        }
    }

    /**
     * 将数据写入模板文件
     */
    // 在 writeDataToTemplate 方法开头设置当前日志回调
    public static void writeDataToTemplate(List<DataRow> dataList,
                                           Map<String, Integer> palletTotals,
                                           String templatePath, String outputPath,
                                           LogCallback logCallback) throws IOException {
        // 设置当前日志回调
        currentLogCallback = logCallback;
        // 读取模板文件
        try (FileInputStream fis = new FileInputStream(templatePath);
             Workbook workbook = WorkbookFactory.create(fis);
             FileOutputStream fos = new FileOutputStream(outputPath)) {

            Sheet sheet = workbook.getSheetAt(0);

            // 1. 确保第15行存在并设置正确的样式
            System.out.println("   准备第15行模板样式...");
            Row templateRow = ensureTemplateRowExists(sheet, TARGET_START_ROW, workbook);

            // 计算需要清空的行数（根据实际数据量）
            int dataRowCount = dataList.size();
            clearDataArea(sheet, TARGET_START_ROW, TARGET_START_ROW + dataRowCount - 1);


            // 3. 开始填充数据
            System.out.println("   填充数据...");

            // 创建各种样式
            CellStyle textStyle = createTextStyle(workbook);  // 文本样式，用于B、I、J、N列
            CellStyle numberStyle = createNumberStyle(workbook); // 数字样式，用于F、H列
            CellStyle integerStyle = createIntegerStyle(workbook); // 整数样式，用于G列
            CellStyle fourDecimalStyle = createFourDecimalStyle(workbook); // 四位小数样式，用于O列
            CellStyle centeredStyle = createCenteredStyle(workbook); // 居中样式，用于其他文本列

            // 样式映射
            Map<Integer, CellStyle> styleMap = new HashMap<>();

            // 为每列设置样式
            for (int col = 0; col <= TGT_COL_PALLETS; col++) {
                if (col == TGT_COL_PALLET_NO || col == TGT_COL_CSG_CODE || col == TGT_COL_CSP_CODE ||
                        col == TGT_COL_TOTAL_CASES_PALLET) {
                    // B、I、J、N列：文本样式
                    styleMap.put(col, textStyle);
                } else if (col == TGT_COL_NW || col == TGT_COL_TOTAL_NW) {
                    // F、H列：数字样式（两位小数）
                    styleMap.put(col, numberStyle);
                } else if (col == TGT_COL_CASES) {
                    // G列：整数样式
                    styleMap.put(col, integerStyle);
                } else if (col == TGT_COL_PALLETS) {
                    // O列：四位小数样式
                    styleMap.put(col, fourDecimalStyle);
                } else {
                    // 其他列：居中样式
                    styleMap.put(col, centeredStyle);
                }
            }

            for (int i = 0; i < dataList.size(); i++) {
                DataRow data = dataList.get(i);
                int currentRowNum = TARGET_START_ROW + i;

                Row row;
                if (i == 0) {
                    // 第一行使用模板行
                    row = templateRow;
                } else {
                    // 创建新行
                    row = sheet.createRow(currentRowNum);

                    // 复制行高
                    row.setHeight(templateRow.getHeight());

                    // 为每个单元格应用样式
                    for (int col = 0; col <= TGT_COL_PALLETS; col++) {
                        Cell newCell = row.createCell(col);
                        CellStyle style = styleMap.get(col);
                        if (style != null) {
                            newCell.setCellStyle(style);
                        }
                    }
                }

                // 填充数据
                fillRowData(row, data, palletTotals, currentRowNum + 1, styleMap); // Excel行号从1开始

                // 显示进度
                if ((i + 1) % 50 == 0 || i == dataList.size() - 1) {
                    System.out.println("   已填充 " + (i + 1) + "/" + dataList.size() + " 行");
                }
            }

            // 新增：处理每行的计算列（替代公式）
            for (int i = 0; i < dataList.size(); i++) {
                int currentRowNum = TARGET_START_ROW + i;
                Row row = sheet.getRow(currentRowNum);

                if (row != null) {
                    // 获取F列(N.W)和G列(CASES)的值
                    Cell cellF = row.getCell(TGT_COL_NW);
                    Cell cellG = row.getCell(TGT_COL_CASES);

                    double nwValue = 0.0;
                    int casesValue = 0;

                    if (cellF != null) {
                        nwValue = cellF.getCellType() == CellType.NUMERIC ? cellF.getNumericCellValue() : 0.0;
                    }
                    if (cellG != null) {
                        casesValue = (int) (cellG.getCellType() == CellType.NUMERIC ? cellG.getNumericCellValue() : 0.0);
                    }

                    // H列: TOTAL N.W = F列 * G列
                    Cell cellH = row.getCell(TGT_COL_TOTAL_NW);
                    if (cellH == null) {
                        cellH = row.createCell(TGT_COL_TOTAL_NW);
                        if (styleMap.get(TGT_COL_TOTAL_NW) != null) {
                            cellH.setCellStyle(styleMap.get(TGT_COL_TOTAL_NW));
                        }
                    }
                    cellH.setCellValue(nwValue * casesValue);

                    // O列: Pallets = G列 / N列
                    Cell cellN = row.getCell(TGT_COL_TOTAL_CASES_PALLET);
                    int totalCasesValue = 0;
                    if (cellN != null) {
                        if (cellN.getCellType() == CellType.STRING) {
                            try {
                                String stringValue = cellN.getStringCellValue();
                                // 移除可能的空格和非数字字符，只保留数字
                                stringValue = stringValue.replaceAll("[^0-9]", "");
                                if (!stringValue.isEmpty()) {
                                    totalCasesValue = Integer.parseInt(stringValue);
                                }
                            } catch (NumberFormatException e) {
                                //System.err.println("解析N列字符串为整数时出错: " + e.getMessage());
                                // 替换 System.err.println 为日志回调
                                if (currentLogCallback != null) {
                                    currentLogCallback.logError("解析N列字符串为整数时出错: " + e.getMessage());
                                }
                            }
                        } else if (cellN.getCellType() == CellType.NUMERIC) {
                            totalCasesValue = (int) cellN.getNumericCellValue();
                        }
                    }


                    Cell cellO = row.getCell(TGT_COL_PALLETS);
                    if (cellO == null) {
                        cellO = row.createCell(TGT_COL_PALLETS);
                        if (styleMap.get(TGT_COL_PALLETS) != null) {
                            cellO.setCellStyle(styleMap.get(TGT_COL_PALLETS));
                        }
                    }

                    if (totalCasesValue != 0) {
                        cellO.setCellValue((double) casesValue / totalCasesValue);
                    } else {
                        cellO.setCellValue(0.0);
                    }
                }
            }

            // 计算汇总值
            double totalCases = 0.0;
            double totalNetKg = 0.0;
            double totalPallets = 0.0;

            for (int i = 0; i < dataList.size(); i++) {
                int currentRowNum = TARGET_START_ROW + i;
                Row row = sheet.getRow(currentRowNum);

                if (row != null) {
                    // 累加G列(CASES)
                    Cell cellG = row.getCell(TGT_COL_CASES);
                    if (cellG != null && cellG.getCellType() == CellType.NUMERIC) {
                        totalCases += cellG.getNumericCellValue();
                    }

                    // 累加H列(TOTAL N.W)
                    Cell cellH = row.getCell(TGT_COL_TOTAL_NW);
                    if (cellH != null && cellH.getCellType() == CellType.FORMULA) {
                        totalNetKg += cellH.getNumericCellValue();
                    }

                    // 累加O列(Pallets)
                    Cell cellO = row.getCell(TGT_COL_PALLETS);
                    if (cellO != null && cellO.getCellType() == CellType.FORMULA) {
                        totalPallets += cellO.getNumericCellValue();
                    }
                }
            }

            // 新增：在第12行(索引11)填充汇总数据
            Row summaryRow = sheet.getRow(11); // 第12行
            if (summaryRow == null) {
                summaryRow = sheet.createRow(11);
            }

            // 填充汇总数据
            // M列: Cases总和
            Cell casesSummaryCell = summaryRow.getCell(12); // M列
            if (casesSummaryCell == null) {
                casesSummaryCell = summaryRow.createCell(12);
            }
            casesSummaryCell.setCellValue(totalCases);
            System.out.println("   M列: Cases总和 = " + totalCases);

            // N列: Net Kg总和
            Cell netKgSummaryCell = summaryRow.getCell(13); // N列
            if (netKgSummaryCell == null) {
                netKgSummaryCell = summaryRow.createCell(13);
            }
            netKgSummaryCell.setCellValue(totalNetKg);
            System.out.println("   N列: Net Kg总和 = " + totalNetKg);

            // O列: Pallets总和
            Cell palletsSummaryCell = summaryRow.getCell(14); // O列
            if (palletsSummaryCell == null) {
                palletsSummaryCell = summaryRow.createCell(14);
            }
            palletsSummaryCell.setCellValue(totalPallets);
            System.out.println("   O列: Pallets总和 = " + totalPallets);

            // 保存工作簿
            workbook.write(fos);
            System.out.println("   数据填充完成！");

        } catch (FileNotFoundException e) {
            System.err.println("模板文件未找到: " + templatePath);
            throw e;
        }
    }

    /**
     * 确保模板行存在并设置正确的样式
     */
    private static Row ensureTemplateRowExists(Sheet sheet, int rowIndex, Workbook workbook) {
        Row row = sheet.getRow(rowIndex);
        if (row == null) {
            row = sheet.createRow(rowIndex);
        }

        // 创建各种样式
        CellStyle textStyle = createTextStyle(workbook);
        CellStyle numberStyle = createNumberStyle(workbook);
        CellStyle integerStyle = createIntegerStyle(workbook);
        CellStyle fourDecimalStyle = createFourDecimalStyle(workbook);
        CellStyle centeredStyle = createCenteredStyle(workbook);

        // 为所有单元格应用样式
        for (int col = 0; col <= TGT_COL_PALLETS; col++) {
            Cell cell = row.getCell(col);
            if (cell == null) {
                cell = row.createCell(col);
            }

            // 根据列类型设置样式
            if (col == TGT_COL_PALLET_NO || col == TGT_COL_CSG_CODE || col == TGT_COL_CSP_CODE ||
                    col == TGT_COL_TOTAL_CASES_PALLET) {
                // B、I、J、N列：文本样式
                cell.setCellStyle(textStyle);
            } else if (col == TGT_COL_NW || col == TGT_COL_TOTAL_NW) {
                // F、H列：数字样式（两位小数）
                cell.setCellStyle(numberStyle);
            } else if (col == TGT_COL_CASES) {
                // G列：整数样式
                cell.setCellStyle(integerStyle);
            } else if (col == TGT_COL_PALLETS) {
                // O列：四位小数样式
                cell.setCellStyle(fourDecimalStyle);
            } else {
                // 其他列：居中样式
                cell.setCellStyle(centeredStyle);
            }
        }

        return row;
    }

    /**
     * 创建居中样式
     */
    private static CellStyle createCenteredStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();

        // 设置水平和垂直居中
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);

        // 设置边框
        style.setBorderTop(BorderStyle.THIN);
        style.setBorderBottom(BorderStyle.THIN);
        style.setBorderLeft(BorderStyle.THIN);
        style.setBorderRight(BorderStyle.THIN);

        // 设置字体
        Font font = workbook.createFont();
        font.setFontHeightInPoints((short) 10);
        style.setFont(font);

        return style;
    }

    /**
     * 创建数字样式（两位小数）
     */
    private static CellStyle createNumberStyle(Workbook workbook) {
        CellStyle style = createCenteredStyle(workbook);

        // 设置数字格式：保留两位小数
        style.setDataFormat(workbook.createDataFormat().getFormat("0.00"));

        return style;
    }

    /**
     * 创建整数样式
     */
    private static CellStyle createIntegerStyle(Workbook workbook) {
        CellStyle style = createCenteredStyle(workbook);

        // 设置整数格式（不要.00后缀）
        style.setDataFormat(workbook.createDataFormat().getFormat("0"));

        return style;
    }

    /**
     * 创建四位小数样式
     */
    private static CellStyle createFourDecimalStyle(Workbook workbook) {
        CellStyle style = createCenteredStyle(workbook);

        // 设置数字格式：保留四位小数
        style.setDataFormat(workbook.createDataFormat().getFormat("0.0000"));

        return style;
    }

    /**
     * 创建文本样式
     */
    private static CellStyle createTextStyle(Workbook workbook) {
        CellStyle style = createCenteredStyle(workbook);

        // 设置文本格式（避免数字显示为.00）
        style.setDataFormat(workbook.createDataFormat().getFormat("@"));

        return style;
    }

    /**
     * 清空数据区域（指定范围）
     */
    private static void clearDataArea(Sheet sheet, int startRow, int endRow) {
        // 只清空从startRow到endRow的行
        for (int rowNum = startRow; rowNum <= Math.min(endRow, sheet.getLastRowNum()); rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row != null) {
                for (int col = 0; col <= TGT_COL_PALLETS; col++) {
                    Cell cell = row.getCell(col);
                    if (cell != null) {
                        cell.setCellValue("");
                    }
                }
            }
        }
    }

    /**
     * 填充行数据
     */
    private static void fillRowData(Row row, DataRow data,
                                    Map<String, Integer> palletTotals, int excelRowNum,
                                    Map<Integer, CellStyle> styleMap) {
        try {
            // A列: Species（固定值"CEREZAS"）
            setCellValue(row, TGT_COL_SPECIES, "CEREZAS");

            // B列: Pallet No.（文本类型）
            Cell cellB = row.getCell(TGT_COL_PALLET_NO);
            if (cellB == null) {
                cellB = row.createCell(TGT_COL_PALLET_NO);
                if (styleMap.get(TGT_COL_PALLET_NO) != null) {
                    cellB.setCellStyle(styleMap.get(TGT_COL_PALLET_NO));
                }
            }
            // 设置文本类型的值（即使数字也按文本处理）
            cellB.setCellValue(data.idPallet);

            // C列: Brand
            setCellValue(row, TGT_COL_BRAND, data.label);

            // D列: Variety
            setCellValue(row, TGT_COL_VARIETY, data.variety);

            // E列: Size
            setCellValue(row, TGT_COL_SIZE, data.size);

            // F列: N.W (数值，两位小数)
            double netWeight = parseNetWeight(data.netWeight);
            setNumericCellValue(row, TGT_COL_NW, netWeight);

            // G列: CASES (整数，不要.00后缀)
            int quantity = parseQuantity(data.quantity);
            Cell cellG = row.getCell(TGT_COL_CASES);
            if (cellG == null) {
                cellG = row.createCell(TGT_COL_CASES);
                if (styleMap.get(TGT_COL_CASES) != null) {
                    cellG.setCellStyle(styleMap.get(TGT_COL_CASES));
                }
            }
            cellG.setCellValue(quantity);

            // H列: TOTAL N.W (公式: F列 * G列)
            Cell cellH = row.getCell(TGT_COL_TOTAL_NW);
            if (cellH == null) {
                cellH = row.createCell(TGT_COL_TOTAL_NW);
                if (styleMap.get(TGT_COL_TOTAL_NW) != null) {
                    cellH.setCellStyle(styleMap.get(TGT_COL_TOTAL_NW));
                }
            }
            String formulaH = "F" + excelRowNum + "*G" + excelRowNum;
            cellH.setCellFormula(formulaH);

            // I列: CSG Code（文本类型）
            Cell cellI = row.getCell(TGT_COL_CSG_CODE);
            if (cellI == null) {
                cellI = row.createCell(TGT_COL_CSG_CODE);
                if (styleMap.get(TGT_COL_CSG_CODE) != null) {
                    cellI.setCellStyle(styleMap.get(TGT_COL_CSG_CODE));
                }
            }
            // 设置文本类型的值，如果csg值含有.00，去掉
            data.csg = data.csg.replace(".00", "");
            cellI.setCellValue(data.csg);

            // J列: CSP Code（文本类型）
            Cell cellJ = row.getCell(TGT_COL_CSP_CODE);
            if (cellJ == null) {
                cellJ = row.createCell(TGT_COL_CSP_CODE);
                if (styleMap.get(TGT_COL_CSP_CODE) != null) {
                    cellJ.setCellStyle(styleMap.get(TGT_COL_CSP_CODE));
                }
            }
            // 设置文本类型的值
            cellJ.setCellValue(data.csp);

            // K列: Packing Date
            setCellValue(row, TGT_COL_PACKING_DATE, data.packingDate);

            // L列: Category
            setCellValue(row, TGT_COL_CATEGORY, data.cat);

            // M列: Thermograph
            setCellValue(row, TGT_COL_THERMOGRAPH, data.tempRecorder);

            // N列: Total cases per pallet (文本类型)
            int totalCasesForPallet = palletTotals.getOrDefault(data.idPallet, 0);
            Cell cellN = row.getCell(TGT_COL_TOTAL_CASES_PALLET);
            if (cellN == null) {
                cellN = row.createCell(TGT_COL_TOTAL_CASES_PALLET);
                if (styleMap.get(TGT_COL_TOTAL_CASES_PALLET) != null) {
                    cellN.setCellStyle(styleMap.get(TGT_COL_TOTAL_CASES_PALLET));
                }
            }
            // 设置文本类型的值
            cellN.setCellValue(String.valueOf(totalCasesForPallet));

            // O列: Pallets (公式: G列 / N列，最多保留4位小数)
            Cell cellO = row.getCell(TGT_COL_PALLETS);
            if (cellO == null) {
                cellO = row.createCell(TGT_COL_PALLETS);
                if (styleMap.get(TGT_COL_PALLETS) != null) {
                    cellO.setCellStyle(styleMap.get(TGT_COL_PALLETS));
                }
            }

            // 设置公式，注意处理除以零的情况
            if (totalCasesForPallet != 0) {
                String formulaO = "G" + excelRowNum + "/N" + excelRowNum;
                cellO.setCellFormula(formulaO);
            } else {
                // 如果总数为0，设置一个简单的公式避免除以零错误
                cellO.setCellFormula("0");
            }

            if (currentLogCallback != null) {
                currentLogCallback.logMessage("   行" + excelRowNum + ": " + data.idPallet +
                        " | 数量: " + quantity +
                        " | 托盘总数: " + totalCasesForPallet);
            }

        } catch (Exception e) {
            // 替换 System.err.println 为日志回调
            if (currentLogCallback != null) {
                currentLogCallback.logError("填充行数据时出错: " + e.getMessage());
                e.printStackTrace();
            }
        }
    }


    /**
     * 设置单元格文本值
     */
    private static void setCellValue(Row row, int colIndex, String value) {
        if (value == null) value = "";

        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }

    /**
     * 设置单元格数值
     */
    private static void setNumericCellValue(Row row, int colIndex, double value) {
        Cell cell = row.getCell(colIndex);
        if (cell == null) {
            cell = row.createCell(colIndex);
        }
        cell.setCellValue(value);
    }

    /**
     * 获取单元格的值
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue().trim();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    return new java.text.SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue());
                } else {
                    double num = cell.getNumericCellValue();
                    if (num == Math.floor(num) && num < 1000000) {
                        return String.valueOf((int) num);
                    } else {
                        return String.format("%.2f", num);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                try {
                    return cell.getStringCellValue();
                } catch (Exception e) {
                    try {
                        return String.valueOf(cell.getNumericCellValue());
                    } catch (Exception ex) {
                        return cell.getCellFormula();
                    }
                }
            default:
                return "";
        }
    }

    /**
     * 检查行是否为空
     */
    private static boolean isRowEmpty(Row row) {
        if (row == null) {
            return true;
        }

        for (int i = 0; i <= row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell != null && cell.getCellType() != CellType.BLANK) {
                String value = getCellValue(cell);
                if (value != null && !value.trim().isEmpty()) {
                    return false;
                }
            }
        }
        return true;
    }
}

/**
 * 数据行类（内部类）
 */
class DataRow {
    String idPallet;
    String label;
    String variety;
    String size;
    String netWeight;
    String quantity;
    String csg;
    String csp;
    String packingDate;
    String cat;
    String tempRecorder;

    @Override
    public String toString() {
        return "DataRow{" +
                "idPallet='" + idPallet + '\'' +
                ", variety='" + variety + '\'' +
                ", size='" + size + '\'' +
                ", quantity='" + quantity + '\'' +
                '}';
    }
}