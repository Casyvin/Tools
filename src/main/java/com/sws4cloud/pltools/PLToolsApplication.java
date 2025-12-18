package com.sws4cloud.pltools;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;
import java.net.URL;

public class PLToolsApplication extends Application {
    @Override
    public void start(Stage stage) throws IOException {
        // 使用更明确的资源加载方式
        URL fxmlLocation = PLToolsApplication.class.getResource("directory-selector.fxml");
        if (fxmlLocation == null) {
            fxmlLocation = PLToolsApplication.class.getResource("/com/sws4cloud/pltools/directory-selector.fxml");
        }

        if (fxmlLocation == null) {
            throw new IOException("Cannot locate directory-selector.fxml");
        }

        FXMLLoader fxmlLoader = new FXMLLoader(fxmlLocation);
        Scene scene = new Scene(fxmlLoader.load(), 800, 600);
        stage.setTitle("PL Tools");
        stage.setScene(scene);
        stage.show();
    }


    public static void main(String[] args) {
        launch();
    }
}