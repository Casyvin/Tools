package com.sws4cloud.pltools;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;
import java.net.URL;

public class PLToolsApplication extends Application {
    @Override
    public void start(Stage stage) {
        try {
            // 确保FXML加载正确
            System.out.println("Application starting...");
            FXMLLoader loader = new FXMLLoader(getClass().getResource("directory-selector.fxml"));
            Parent root = loader.load();

            Scene scene = new Scene(root, 800, 600);
            stage.setTitle("PL Tools");
            stage.setScene(scene);
            stage.show();
        } catch (Exception e) {
            e.printStackTrace(); // 打印错误信息
            System.exit(1);
        }
    }

    public static void main(String[] args) {
        launch(args);
    }
}
