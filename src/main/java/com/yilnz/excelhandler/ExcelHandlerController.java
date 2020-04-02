package com.yilnz.excelhandler;

import javafx.collections.ObservableList;
import javafx.event.EventHandler;
import javafx.fxml.FXML;
import javafx.fxml.Initializable;
import javafx.scene.control.*;
import javafx.scene.input.ContextMenuEvent;
import javafx.stage.DirectoryChooser;
import javafx.stage.FileChooser;
import org.apache.commons.lang3.StringUtils;

import java.io.*;
import java.net.URL;
import java.nio.file.*;
import java.util.*;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Collectors;

public class ExcelHandlerController implements Initializable {

	@FXML
	private Button btnBrowse;
	@FXML
	private Button btnBrowse2;
	@FXML
	private Button btn1;
	@FXML
	private Button btn2;
	@FXML
	private Button btn3;
	@FXML
	private Button btn4;
	@FXML
	private TextField input1;
	@FXML
	private TextField input2;
	@FXML
	private Button btn5;
	@FXML
	private Button btnCSV;
	@FXML
	private TextField inputCSV;
	@FXML
	private ListView<Line> listCSV;


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
		btnBrowse2.setOnAction(e -> {
			final DirectoryChooser fileChooser = new DirectoryChooser();
			fileChooser.setTitle("Open Resource File");
			final File file = fileChooser.showDialog(ExcelHandlerMain.primaryStage);
			if (file != null) {
				input1.setText(file.getPath());
			}
		});
		btn1.setOnAction(e -> {
			if (validate()) return;
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
		btn2.setOnAction(e -> {
			if (validate()) return;
			try {
				new ExcelHandler().handleExcelSeq2(new File(input1.getText()));
				final Alert alert = new Alert(Alert.AlertType.INFORMATION);
				alert.setContentText("处理完毕, 已生成新的文件");
				alert.showAndWait();
			} catch (Exception ee) {
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ee.getMessage());
				alert.showAndWait();
			}
		});
		btn3.setOnAction(e -> {
			if (validate()) return;
			try {
				new ExcelHandler().handleExcelSeq3(new File(input1.getText()), input2.getText());
				final Alert alert = new Alert(Alert.AlertType.INFORMATION);
				alert.setContentText("处理完毕, 已生成新的文件");
				alert.showAndWait();
			} catch (Exception ee) {
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ee.getMessage());
				alert.showAndWait();
			}
		});
		btn4.setOnAction(e -> {
			if (validate()) return;
			try {
				final Alert alert = new Alert(Alert.AlertType.INFORMATION);
				alert.setContentText(new ExcelHandler().checkExcelBlankRow(new File(input1.getText())));
				alert.showAndWait();
			} catch (Exception ee) {
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ee.getMessage());
				alert.showAndWait();
			}
		});
		btn5.setOnAction(e -> {
			if (validate("请选择文件夹")) return;
			try {
				final Alert alert = new Alert(Alert.AlertType.INFORMATION);
				final File file = new File(input1.getText());
				List<File> collect = null;
				if (file.isDirectory()) {
					collect = Arrays.stream(file.listFiles()).filter(f -> f.getName().endsWith(".xlsx") || f.getName().endsWith(".xls")).collect(Collectors.toList());
				} else {
					collect = Arrays.asList(file);
				}

				new ExcelHandler2().handleExcelDelFirstGroupLine(collect);
				alert.setContentText("处理完毕。");
				alert.showAndWait();
			} catch (Exception ee) {
				ee.printStackTrace();
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ee.getClass().toString() + ":" + ee.getMessage());
				alert.showAndWait();
			}
		});
		listCSV.getSelectionModel().setSelectionMode(SelectionMode.MULTIPLE);
		MenuItem menuItem = new MenuItem("创建文件夹");
		listCSV.setContextMenu(new ContextMenu(menuItem));
		menuItem.setOnAction(e -> {
			try {
				ObservableList<Line> selectedItems = listCSV.getSelectionModel().getSelectedItems();
				String parent = new File(inputCSV.getText()).getParent();
				File newDir = new File(parent, selectedItems.get(0).pin);
				if (!newDir.exists()) {
					newDir.mkdir();
				}
				for (Line selectedItem : selectedItems) {
					Path source = Paths.get(parent, selectedItem.dir);
					Path target = Paths.get(new File(newDir, selectedItem.dir).toURI());
					Files.copy(source, target, StandardCopyOption.REPLACE_EXISTING);
					File[] files = source.toFile().listFiles();
					for (File file : files) {
						Files.copy(Paths.get(file.toURI()), Paths.get(target.toString(), file.getName()),  StandardCopyOption.REPLACE_EXISTING);
					}
				}
			} catch (IOException ex) {
				ex.printStackTrace();
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ex.getClass().toString() + ":" + ex.getMessage());
				alert.showAndWait();
			}
		});
		btnCSV.setOnAction(e -> {
			try {
				String filePath = inputCSV.getText();
				BufferedReader br = new BufferedReader(new InputStreamReader(new FileInputStream(new File(filePath))));
				String str = null;
				List<String> lineList = new ArrayList<>();
				while ((str = br.readLine()) != null) {
					lineList.add(str);
				}

				List<Line> list = new ArrayList<>();
				for (String line : lineList) {
					String[] split = line.split(",");
					String s001 = split[1];
					Matcher m = Pattern.compile("S(\\d{4})").matcher(s001);
					if (m.find()) {
						String group = m.group(1);
						Line line1 = new Line();
						line1.dir = split[0];
						line1.number = Integer.parseInt(group);
						line1.pin = s001;
						list.add(line1);
					}
				}
				list.sort(new Comparator<Line>() {
					@Override
					public int compare(Line o1, Line o2) {
						return o1.number - o2.number;
					}
				});
				for (Line line : list) {
					listCSV.getItems().add(line);
				}
			} catch (IOException ex) {
				ex.printStackTrace();
				final Alert alert = new Alert(Alert.AlertType.ERROR);
				alert.setContentText(ex.getClass().toString() + ":" + ex.getMessage());
				alert.showAndWait();
			}
		});
	}

	private boolean validate() {
		return validate("请选择EXCEL文件");
	}

	private boolean validate(String msg) {
		if (StringUtils.isBlank(input1.getText())) {
			final Alert alert = new Alert(Alert.AlertType.ERROR);
			alert.setContentText(msg);
			alert.showAndWait();
			return true;
		}
		return false;
	}
}
