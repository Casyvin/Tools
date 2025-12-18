module com.sws4cloud.pltools {
    requires javafx.controls;
    requires javafx.fxml;
    requires javafx.web;

    requires org.controlsfx.controls;
    requires com.dlsc.formsfx;
    requires net.synedra.validatorfx;
    requires org.kordamp.bootstrapfx.core;
    requires eu.hansolo.tilesfx;
    requires com.almasb.fxgl.all;

    opens com.sws4cloud.pltools to javafx.fxml;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;

    exports com.sws4cloud.pltools;
}