package com.yilnz.excelhandler;

import javafx.application.Application;
import javafx.fxml.FXMLLoader;
import javafx.scene.Parent;
import javafx.scene.Scene;
import javafx.stage.Stage;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.net.URL;

public class ExcelHandlerMain  extends Application {
	private static final Logger logger = LoggerFactory.getLogger(ExcelHandlerMain.class);

	public static Stage primaryStage;

	@Override
	public void start(Stage primaryStage) {
		try {
			ExcelHandlerMain.primaryStage = primaryStage;
			final URL resource = this.getClass().getResource("/main.fxml");
			final Parent main = FXMLLoader.load(resource);
			final Scene scene = new Scene(main);
			primaryStage.setScene(scene);
			primaryStage.show();
		} catch (IOException e) {
			logger.error("primary error", e);
		}

	}
}
