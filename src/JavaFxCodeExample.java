import java.io.File;

import javafx.application.Application;
import javafx.event.ActionEvent;
import javafx.event.EventHandler;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.ComboBox;
import javafx.scene.control.Label;
import javafx.scene.control.TextField;
import javafx.scene.layout.GridPane;
import javafx.stage.FileChooser;
import javafx.stage.Stage;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.Random;

import org.apache.commons.codec.binary.Base64;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class JavaFxCodeExample extends Application {
    File newFile;
    Map<String, Integer> sheets;
    Map<String, Integer> columns;
    GridPane root = new GridPane();

    public static void main(String[] args) {
        launch(args);
    }

    @Override
    public void start(Stage primaryStage) {
        primaryStage.setTitle("Excel Conversion");

        Button fileChoose = new Button();
        fileChoose.setText("Choose file");

        Label encryptionType = new Label("Conversion Type");
        Label selectFile = new Label("Select File");
        Label selectSheet = new Label("Select Sheet");
        Label selectColumn = new Label("Select Column Type");

        TextField fileName = new TextField();
        fileName.setEditable(false);
        TextField tf2 = new TextField();

        Button apply = new Button();
        apply.setText("Apply");

        ComboBox<String> columnNames = new ComboBox<String>();
        columnNames.setValue("Choose");

        ComboBox<String> columnNumbers = new ComboBox<String>();
        columnNumbers.setValue("Choose");

        ComboBox<String> convertType = new ComboBox<String>();
        convertType.getItems().addAll(
                "Encrypt",
                "Decrypt"
        );
        convertType.setValue("Choose");

        ComboBox<String> chooseColumnType = new ComboBox<String>();
        chooseColumnType.getItems().addAll(
                "Column Name",
                "Column Number"
        );

        ComboBox<String> sheetNames = new ComboBox<String>();
        sheetNames.setValue("Choose");

        chooseColumnType.setValue("Choose");
        chooseColumnType.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                if ("Column Number".equals(chooseColumnType.getValue())) {
                    root.getChildren().remove(columnNames);
                    root.getChildren().remove(columnNumbers);
                    root.add(columnNumbers, 2, 3);
                    columns = getColmnNamesOfSheet(sheets.get(sheetNames.getValue()) - 1);
                    removeAllItems(columnNumbers);
                    for (Map.Entry<String, Integer> entry : columns.entrySet()) {
                        columnNumbers.getItems().add(entry.getValue() + "");
                    }
                } else if ("Column Name".equals(chooseColumnType.getValue())) {
                    root.getChildren().remove(columnNumbers);
                    root.getChildren().remove(columnNames);
                    root.add(columnNames, 2, 3);
                    columns = getColmnNamesOfSheet(sheets.get(sheetNames.getValue()) - 1);
                    removeAllItems(columnNames);
                    for (Map.Entry<String, Integer> entry : columns.entrySet()) {
                        columnNames.getItems().add(entry.getKey());
                    }
                }
            }
        });

        columnNames.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {

            }
        });

        fileChoose.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                FileChooser file = new FileChooser();
                file.setTitle("Open File");
                newFile = file.showOpenDialog(primaryStage);
                if (newFile != null) {
                    fileName.setText(newFile.getAbsolutePath());
                    sheets = getSheets();
                    removeAllItems(sheetNames);
                    for (Map.Entry<String, Integer> entry : sheets.entrySet()) {
                        sheetNames.getItems().add(entry.getKey());
                    }
                }
            }
        });

        apply.setOnAction(new EventHandler<ActionEvent>() {

            @Override
            public void handle(ActionEvent event) {
                int index = Integer.MIN_VALUE;
            	boolean temp = "Encrypt".equals(convertType.getValue());

                if ("Column Number".equals(chooseColumnType.getValue())) {
                    index = Integer.parseInt(columnNumbers.getValue()) - 1;
                } else if ("Column Name".equals(chooseColumnType.getValue())) {
                    index = getColumIndex(sheets.get(sheetNames.getValue()), columnNames.getValue(), newFile.getAbsolutePath());
                }
                updateExcelFile(sheets.get(sheetNames.getValue()), index, newFile.getAbsolutePath(), temp);
            }
        });

        root.setVgap(10);
        root.setHgap(10);
        root.setPadding(new Insets(5, 5, 5, 5));

        root.add(encryptionType, 0, 0);
        root.add(convertType, 1, 0);

        root.add(selectFile, 0, 1);
        root.add(fileChoose, 1, 1);
        root.add(fileName, 2, 1);

        root.add(selectSheet, 0, 2);
        root.add(sheetNames, 1, 2);

        root.add(selectColumn, 0, 3);
        root.add(chooseColumnType, 1, 3);

        root.add(apply, 0, 4);

        primaryStage.setScene(new Scene(root, 400, 180));
        primaryStage.setResizable(false);
        primaryStage.show();
    }

    private void removeAllItems(ComboBox list) {
        int size = list.getItems().size();
        for (int i = 0; i < size; i++) {
            list.getItems().remove(0);
        }
        list.setValue("Choose");
    }

    private String handleEncodeScope(String encoded) {
        Random rand = new Random();
        int number = rand.nextInt(10);
        for (int i = 0; i <= number; i++) {
            encoded = encodeString(encoded);
        }
        encoded += number;
        return encoded;
    }

    private String handleDecodeScope(String decode) {
        String decoded = decode.substring(0, decode.length() - 1);
        int num = Integer.parseInt(decode.charAt(decode.length() - 1) + "");
        for (int i = 0; i <= num; i++) {
            decoded = decodeString(decoded);
        }
        return decoded;
    }

    private String encodeString(String str) {
        String encoded = new String(Base64.encodeBase64(str.getBytes()));
        return encoded;
    }

    private String decodeString(String str) {
        String decoded = new String(Base64.decodeBase64(str.getBytes()));
        return decoded;
    }

    private void updateExcelFile(int sheetIndex, int n, String fileName, boolean res) {
        try {
            FileInputStream inputStream = new FileInputStream(new File(fileName));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet newSheet = workbook.getSheetAt(sheetIndex - 1);
            for (int i = 1; i <= newSheet.getLastRowNum(); i++) {
                Row row = newSheet.getRow(i);
                Cell cell = row.getCell(n);
                if (cell != null) {
                    if (res) {
                        String encoded = handleEncodeScope(cell.getStringCellValue());
                        cell.setCellValue(encoded);
                    } else {
                        String decode = handleDecodeScope(cell.getStringCellValue());
                        cell.setCellValue(decode);
                    }
                }
            }
            System.out.println("Updated successfully : ");
            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
    }

    public int getColumIndex(int sheetIndex, String str, String fileName) {
        int result = Integer.MIN_VALUE;
        try {
            FileInputStream inputStream = new FileInputStream(new File(fileName));
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet newSheet = workbook.getSheetAt(sheetIndex - 1);
            Row row = newSheet.getRow(0);
            for (int j = 0; j <= row.getLastCellNum(); j++) {
                Cell cell = row.getCell(j);
                if (str.equals(cell.getStringCellValue())) {
                    result = j;
                    break;
                }
            }
            FileOutputStream outputStream = new FileOutputStream(fileName);
            workbook.write(outputStream);
            workbook.close();
            outputStream.close();

        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return result;
    }

    public Map<String, Integer> getSheets() {
        int result = 0;
        Map<String, Integer> list = new HashMap<>();
        try {
            FileInputStream inputStream = new FileInputStream(newFile);
            Workbook workbook = WorkbookFactory.create(inputStream);
            result = workbook.getNumberOfSheets();
            for (int i = 0; i < result; i++) {
                list.put(workbook.getSheetName(i), i + 1);
            }
            workbook.close();
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return list;
    }

    public Map<String, Integer> getColmnNamesOfSheet(int index) {
        Map<String, Integer> list = new HashMap<>();
        try {
            FileInputStream inputStream = new FileInputStream(newFile);
            Workbook workbook = WorkbookFactory.create(inputStream);
            Sheet newSheet = workbook.getSheetAt(index);
            for (int i = 0; i < 1; i++) {
                Row row = newSheet.getRow(i);
                for (int j = 0; j < row.getLastCellNum(); j++) {
                    Cell cell = row.getCell(j);
                    list.put(cell.getStringCellValue(), j + 1);
                }
            }
        } catch (Exception ex) {
            ex.printStackTrace();
        }
        return list;
    }
}