module com.sws4cloud.pltools {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;
    requires javafx.swing;
    requires javafx.media;
    requires org.apache.poi.poi;

    opens com.sws4cloud.pltools to javafx.fxml;
    exports com.sws4cloud.pltools;
}