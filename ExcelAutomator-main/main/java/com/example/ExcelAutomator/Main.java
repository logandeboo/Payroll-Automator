package com.example.ExcelAutomator;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;

import java.io.IOException;

public class Main extends Application {
    @Override
    public void start(Stage primaryStage) throws IOException {
        FXMLLoader loader = new FXMLLoader(getClass().getResource("Hub.fxml"));
        Parent root = loader.load();
        primaryStage.setTitle("Report Generator");
        primaryStage.setScene(new Scene(root, 900.0, 600.0));
        root.getStylesheets().clear();
        primaryStage.show();
    }

    public static void main(String[] args) {
        launch();
    }
}