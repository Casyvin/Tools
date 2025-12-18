package com.sws4cloud.pltools;

import javafx.fxml.FXML;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.TextArea;
import javafx.scene.control.TextField;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;

import java.io.File;
import java.net.MalformedURLException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;

public class DirectorySelectorController {

    @FXML
    private Label sourceDirLabel;

    @FXML
    private Label targetDirLabel;

    @FXML
    private TextField sourceDirectoryField;

    @FXML
    private TextField targetDirectoryField;

    @FXML
    private Button chooseSourceButton;

    @FXML
    private Button chooseTargetButton;

    @FXML
    private Button executeButton;

    @FXML
    private Button languageToggleButton;

    @FXML
    private Label statusLabel;

    @FXML
    private Label logLabel;

    @FXML
    private TextArea logTextArea;

    private File sourceDirectory;
    private File targetDirectory;
    private boolean isChinese = true; // 默认为中文

    @FXML
    private TextField templateFileField;

    @FXML
    private Button chooseTemplateButton;

    private File templateFile;

    @FXML
    private Label templateFileLabel;

    @FXML
    protected void onChooseTemplateButtonClick() {
        FileChooser fileChooser = new FileChooser();
        fileChooser.setTitle(isChinese ? "选择模板文件" : "Select Template File");
        fileChooser.getExtensionFilters().add(new FileChooser.ExtensionFilter("Excel Files", "*.xlsx"));

        if (templateFile != null && templateFile.exists()) {
            fileChooser.setInitialDirectory(templateFile.getParentFile());
        } else if (sourceDirectory != null && sourceDirectory.exists()) {
            fileChooser.setInitialDirectory(sourceDirectory);
        }

        File selectedFile = fileChooser.showOpenDialog(templateFileField.getScene().getWindow());
        if (selectedFile != null) {
            templateFile = selectedFile;
            templateFileField.setText(selectedFile.getAbsolutePath());
            statusLabel.setText(isChinese ?
                    "已选择模板文件: " + selectedFile.getName() :
                    "Selected template file: " + selectedFile.getName());
        }
    }

    @FXML
    protected void onExecuteButtonClick() throws MalformedURLException {
        // 检查是否已选择必要目录
        if (getSourceDirectory() == null || getTargetDirectory() == null) {
            statusLabel.setText(isChinese ? "请先选择源目录和目标目录！" : "Please select source and target directories first!");
            appendLog(isChinese ? "错误: 请先选择源目录和目标目录！" : "Error: Please select source and target directories first!");
            return;
        }

        // 检查是否已选择模板文件
        String templateFilePath;
        if (templateFile != null && templateFile.exists()) {
            templateFilePath = templateFile.getAbsolutePath();
            appendLog(isChinese ? "使用用户选择的模板文件: " + templateFilePath :
                    "Using user selected template file: " + templateFilePath);
        } else {
            // 使用默认模板文件加载逻辑
            String templateResourcePath = "templates/PL-Template1.xlsx";
            ClassLoader classLoader = Thread.currentThread().getContextClassLoader();
            java.net.URL resourceUrl = classLoader.getResource(templateResourcePath);

            if (resourceUrl == null) {
                // 如果类路径加载失败，尝试从文件系统加载
                String projectPath = System.getProperty("user.dir");
                File defaultTemplateFile = new File(projectPath + "/src/main/resources/templates/PL-Template1.xlsx");

                if (defaultTemplateFile.exists()) {
                    templateFilePath = defaultTemplateFile.getAbsolutePath();
                    appendLog(isChinese ? "使用默认模板文件: " + templateFilePath :
                            "Using default template file: " + templateFilePath);
                } else {
                    String errorMsg = isChinese ? "未找到模板文件" : "Template file not found";
                    appendLog("ERROR: " + errorMsg);
                    statusLabel.setText(errorMsg);
                    return;
                }
            } else {
                templateFilePath = resourceUrl.getPath();
                appendLog(isChinese ? "使用内置模板文件" : "Using built-in template file");
            }
        }

        executeButton.setDisable(true);
        appendLog(isChinese ? "开始执行Excel数据提取任务..." : "Starting Excel data extraction task...");

        // 在后台线程中执行耗时操作
        new Thread(() -> {
            try {
                ExcelDataExtractor.executeDataExtraction(
                        templateFilePath,
                        getSourceDirectory().getAbsolutePath(),
                        getTargetDirectory().getAbsolutePath(),
                        new ExcelDataExtractor.LogCallback() {
                            @Override
                            public void logMessage(String message) {
                                // 在JavaFX主线程中更新UI
                                javafx.application.Platform.runLater(() -> appendLog(message));
                            }

                            @Override
                            public void logError(String message) {
                                // 在JavaFX主线程中更新UI
                                javafx.application.Platform.runLater(() -> {
                                    appendLog("ERROR: " + message);
                                    statusLabel.setText(isChinese ? "执行出错: " + message : "Execution error: " + message);
                                });
                            }
                        }
                );
            } finally {
                // 重新启用执行按钮
                javafx.application.Platform.runLater(() -> {
                    executeButton.setDisable(false);
                    statusLabel.setText(isChinese ? "任务执行完成" : "Task completed");
                });
            }
        }).start();
    }

    @FXML
    private Button clearLogButton;

    @FXML
    protected void onClearLogButtonClick() {
        logTextArea.clear();
    }

    @FXML
    protected void initialize() {
        // 初始化时设置清除日志按钮的文字（根据当前语言）
        if (isChinese) {
            clearLogButton.setText("清除日志");
        } else {
            clearLogButton.setText("Clear Log");
        }
    }

    @FXML
    protected void onChooseSourceButtonClick() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle(isChinese ? "选择源文件目录" : "Select Source Directory");

        if (sourceDirectory != null && sourceDirectory.exists()) {
            directoryChooser.setInitialDirectory(sourceDirectory);
        }

        File selectedDirectory = directoryChooser.showDialog(
                sourceDirectoryField.getScene().getWindow());

        if (selectedDirectory != null) {
            sourceDirectory = selectedDirectory;
            sourceDirectoryField.setText(selectedDirectory.getAbsolutePath());
            statusLabel.setText(isChinese ?
                    "已选择源目录: " + selectedDirectory.getName() :
                    "Selected source directory: " + selectedDirectory.getName());
        }
    }

    @FXML
    protected void onChooseTargetButtonClick() {
        DirectoryChooser directoryChooser = new DirectoryChooser();
        directoryChooser.setTitle(isChinese ? "选择目标文件目录" : "Select Target Directory");

        if (targetDirectory != null && targetDirectory.exists()) {
            directoryChooser.setInitialDirectory(targetDirectory);
        } else if (sourceDirectory != null && sourceDirectory.exists()) {
            directoryChooser.setInitialDirectory(sourceDirectory);
        }

        File selectedDirectory = directoryChooser.showDialog(
                targetDirectoryField.getScene().getWindow());

        if (selectedDirectory != null) {
            targetDirectory = selectedDirectory;
            targetDirectoryField.setText(selectedDirectory.getAbsolutePath());
            statusLabel.setText(isChinese ?
                    "已选择目标目录: " + selectedDirectory.getName() :
                    "Selected target directory: " + selectedDirectory.getName());
        }
    }

    @FXML
    protected void onLanguageToggleClick() {
        isChinese = !isChinese;
        updateLanguage();
    }

    private void performBusinessLogic() {
        appendLog(isChinese ? "正在扫描源目录..." : "Scanning source directory...");

        if (sourceDirectory.listFiles() != null) {
            int fileCount = sourceDirectory.listFiles().length;
            appendLog((isChinese ? "源目录中共有 " : "Found ") + fileCount +
                    (isChinese ? " 个文件/文件夹" : " files/folders in source directory"));
        }

        appendLog(isChinese ? "文件处理完成" : "File processing completed");
    }

    private void appendLog(String message) {
        String timestamp = LocalDateTime.now().format(DateTimeFormatter.ofPattern("HH:mm:ss"));
        logTextArea.appendText("[" + timestamp + "] " + message + "\n");
        logTextArea.positionCaret(logTextArea.getText().length());
    }

    public File getSourceDirectory() {
        return sourceDirectory;
    }

    public File getTargetDirectory() {
        return targetDirectory;
    }

    private void updateLanguage() {
        if (isChinese) {
            sourceDirLabel.setText("源文件目录:");
            targetDirLabel.setText("结果存储目录:");
            templateFileLabel.setText("模板文件:");
            chooseSourceButton.setText("浏览...");
            chooseTargetButton.setText("浏览...");
            chooseTemplateButton.setText("浏览...");
            executeButton.setText("执行");
            languageToggleButton.setText("English");
            statusLabel.setText("请选择目录");
            logLabel.setText("执行日志:");
            clearLogButton.setText("清除日志");
        } else {
            sourceDirLabel.setText("Source Directory:");
            targetDirLabel.setText("Result Directory:");
            templateFileLabel.setText("Template File:");
            chooseSourceButton.setText("Browse...");
            chooseTargetButton.setText("Browse...");
            chooseTemplateButton.setText("Browse...");
            executeButton.setText("Execute");
            languageToggleButton.setText("中文");
            statusLabel.setText("Please select directories");
            logLabel.setText("Execution Log:");
            clearLogButton.setText("Clear Log");
        }
    }


}
