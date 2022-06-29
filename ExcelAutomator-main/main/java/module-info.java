module com.example.testproject {
    requires javafx.controls;
    requires javafx.fxml;
    requires java.sql;
    requires java.datatransfer;
    requires org.apache.poi.poi;
    requires org.apache.poi.ooxml;
    requires java.desktop;


    opens com.example.ExcelAutomator to javafx.fxml;
    exports com.example.ExcelAutomator;
}