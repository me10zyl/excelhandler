package com.yilnz.excelhandler;

import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.Alert;
import javafx.scene.control.Button;
import javafx.scene.control.TextField;
import javafx.stage.FileChooser;
import org.apache.commons.lang3.StringUtils;

import java.io.File;
import java.net.URL;
import java.util.ResourceBundle;

public class ExcelHandlerController implements Initializable {

	@FXML
	private Button btnBrowse;
	@FXML
	private Button btnStart;
	@FXML
	private TextField input1;

	@Override
	public void initialize(URL location, ResourceBundle resources) {
		btnBrowse.setOnAction(e -> {
			final FileChooser fileChooser = new FileChooser();
			fileChooser.setSelectedExtensionFilter(new FileChooser.ExtensionFilter("EXCEL文件", "xlsx", "xls"));
			fileChooser.setTitle("Open Resource File");
			final File file = fileChooser.showOpenDialog(ExcelHandlerMain.primaryStage);
			if (file != null) {
				input1.setText(file.getPath());
			}
		});
		btnStart.setOnAction(e -> {
			if (StringUtils.isBlank(input1.getText())) {
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText("请选择EXCEL文件");
				alert.showAndWait();
				return;
			}
			try {
				new ExcelHandler().handleExcelSeq(new File(input1.getText()));
				final Alert alert = new Alert(Alert.AlertType.INFORMATION);
				alert.setContentText("处理完毕, 已生成新的文件");
				alert.showAndWait();
			} catch (Exception ee) {
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ee.getMessage());
				alert.showAndWait();
			}
		});
	}
}
