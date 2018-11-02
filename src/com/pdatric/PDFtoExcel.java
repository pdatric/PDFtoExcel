/*
 * Made by Patric Luebbert 2016
 * Specifically designed for ICP-MS PDF files generated in Brewer Science
 */

package com.pdatric;

import edu.duke.FileResource;
import javafx.application.Application;
import javafx.geometry.HPos;
import javafx.geometry.Insets;
import javafx.scene.Scene;
import javafx.scene.control.Button;
import javafx.scene.control.Label;
import javafx.scene.control.*;
import javafx.scene.control.TextField;
import javafx.scene.layout.ColumnConstraints;
import javafx.scene.layout.GridPane;
import javafx.scene.layout.Priority;
import javafx.scene.layout.RowConstraints;
import javafx.stage.FileChooser;
import javafx.stage.FileChooser.ExtensionFilter;
import javafx.stage.Stage;
import javafx.scene.image.Image;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;

import java.awt.*;
import java.awt.event.KeyAdapter;
import java.io.*;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Alert.AlertType;
import javafx.scene.layout.Region;

import static org.apache.poi.ss.usermodel.CellStyle.ALIGN_CENTER;

/**
 * @author pluebbert
 */
public class PDFtoExcel extends Application {
    private static TextField field;
    private static String outputName;
    private static int size;

    private static boolean isEmpty;

    private ListView<String> listViewStrings;

    private Stage savedStage;

    private static File tmp;
    private static File csvFile;

    private static String OStype; 
    private static String userName;
    private static File dir;
   
    
    
    private static String dirPath;
    private static String csvFilePath;
    private static String tmpPath;
    private static String IonsCopyPath;
     
    
    private static List<String> selStrings = new ArrayList();
    private static List<String> shortStrings = new ArrayList();
    ////////////////////////////
    // Class Functions
    ////////////////////////////

    //Start the GUI
    @Override
    public void start(Stage stage) {
       // stage.getIcons().add(new Image("/Users/pluebbert/NetBeansProjects/PDFtoExcel-alex-revision/Icon"));
        OStype = System.getProperty("os.name");
            System.out.println("OS Type: " + OStype);
        userName = System.getProperty("user.name");
            System.out.println("User Name: " + userName);
        
       String path = (File.listRoots()[0].toString() + System.getProperty("file.separator") + "Users" + System.getProperty("file.separator") + userName + System.getProperty("file.separator") + "Desktop"); 
       
       String dirName = "Temp_Files"; 
        
        File file = new File(path);
        file.mkdirs();
        
        dir = new File(path + System.getProperty("file.separator") + dirName);
        try{
            if(dir.mkdirs()) {
                }
        
        }catch(Exception ex) {
            ex.printStackTrace();
        }
        
        dirPath = (path + System.getProperty("file.separator") + dirName);
            System.out.println(dirPath);
        
        
        GridPane root = new GridPane();
        root.setHgap(8);
        root.setVgap(8);
        root.setPadding(new Insets(5));

        ColumnConstraints cons1 = new ColumnConstraints();
        cons1.setHgrow(Priority.NEVER);
        root.getColumnConstraints().add(cons1);
        ColumnConstraints cons2 = new ColumnConstraints();
        cons2.setHgrow(Priority.ALWAYS);
        root.getColumnConstraints().addAll(cons1, cons2);

        RowConstraints rcons1 = new RowConstraints();
        rcons1.setVgrow(Priority.NEVER);
        RowConstraints rcons2 = new RowConstraints();
        rcons2.setVgrow(Priority.ALWAYS);
        root.getRowConstraints().addAll(rcons1, rcons2);

        // Visuals
        Label lbl = new Label("File Name:");
        Label author = new Label("Made by Patric Luebbert");

        field = new TextField();
        
        listViewStrings = new ListView<>();

        Button okBtn = new Button("Ok");
        Button selectBtn = new Button("Select");
        Button deleteBtn = new Button("Delete");
        Button clearBtn = new Button("Clear");
        Button moveUp = new Button("Move Up");
        Button moveDown = new Button("Move Up");  
        
        
        // Actions
        okBtn.setOnAction(e -> {
            //OK button
            System.out.println("OK");
            outputName = field.getText();  // if desired output file name exists, use it
            
            if (field.getText().isEmpty()) {  //checks if user inputed a desired output file name
                isEmpty = true;             // if true uses default name
                           
            }
            
            
            if(outputName.contains("/")|| outputName.contains("\\")) {
                System.out.println("output name cannot contain a slash");
                
                
                Alert alert = new Alert(AlertType.ERROR, "File Output name CANNOT contain slashes");
                alert.setTitle("File Naming error");
                alert.setHeaderText("File Naming Error");
                alert.getDialogPane().setMinHeight(Region.USE_PREF_SIZE);
                alert.show();
                
                field.getText().replaceAll("\\","_");
                field.getText().replaceAll("/","_");
               
                
                isEmpty = true;
               
            }
            
            // Run Excel Template Creator
            if (selStrings != null) {
                size = selStrings.size();
                excelTemplate();
            }
            
        });

        selectBtn.setOnAction(e -> {
           if(selStrings.isEmpty()){
               System.out.println("No files selected yet");
           
            //Select PDF's Button
            System.out.println("SelectBtn"); //Press to select all ICP-MS PDF files you want to analyze

            //opens file directory to find and select PDF Files
            FileChooser fileChooser = new FileChooser();
            fileChooser.setTitle("Select PDF Files");

            System.out.println("Starting FileChooser");
            
            fileChooser.getExtensionFilters().addAll(new ExtensionFilter("PDF Files", "*.pdf"));
            List<File> selectedFiles = fileChooser.showOpenMultipleDialog(savedStage);

            for (File selectedFile : selectedFiles) {  // gets all of the path's to selected files and saves them as a string
                String tempFilePath = selectedFile.getAbsolutePath();
                System.out.println(tempFilePath);
                selStrings.add(tempFilePath);
                
            }
           }
           else{
               secondarySelecteFile();
           } 

            //shortening up the listview path, need to update list view using this then add a button to increase or decrease the size of path
            System.out.println("Selected String Size: " + selStrings.size());
            for (int i= 0; i < selStrings.size(); i++) {
                String tempFilePath = selStrings.get(i);
               if(OStype.contains("Mac")){
                shortenName();
                   
               }
               if(OStype.contains("Windows")){
                shortStrings.add(selStrings.get(i));
                System.out.println("Windows: " + selStrings);
                
               }
             
            }
            
            Collections.sort(selStrings);
            Collections.sort(shortStrings);//sorts alphabetically for initial view in listview
            refreshListView();
        });

        deleteBtn.setOnAction(e -> {
            //Delete button
            System.out.println("Delete");

            System.out.println("Path Separator: " + System.getProperty("path.separator"));
            System.out.println("line Separator: " + System.getProperty("line.separator"));
            System.out.println("OS Arch: " + System.getProperty("os.arch"));
            System.out.println("File Separator: " + System.getProperty("file.separator"));
            System.out.println("OS name: " + System.getProperty("os.name"));
            System.out.println("OS version: " + System.getProperty("os.version"));
           
            int selectedIndex = listViewStrings.getSelectionModel().getSelectedIndex();
            listViewStrings.getSelectionModel().clearSelection();
            selStrings.remove(selectedIndex);
            shortStrings.remove(selectedIndex);
            refreshListView();
        });

        moveUp.setOnAction(e -> {
            //Move selected Up
            System.out.println("Move up");
            int selectedIndex = listViewStrings.getSelectionModel().getSelectedIndex();
            if (selectedIndex == 0) {
                System.out.println("Already at top of list");
            } else {
                moveUp(selectedIndex);
            }
        });

        moveDown.setOnAction(e -> {
            //Move selected down
            System.out.println("Move Down");
            int selectedIndex = listViewStrings.getSelectionModel().getSelectedIndex();
            if (selectedIndex == selStrings.size()) {
                System.out.println("Already at bottom of list");
            } else {
                moveDown(selectedIndex);
            }
        });

        // Tooltips
        field.setTooltip(new Tooltip("Type what you would like the resulting .xls to be named."
                + " If left blank, the name will be the name and lot# of the sample"));
        clearBtn.setTooltip(new Tooltip("Clears the program for a new workbook"));
        okBtn.setTooltip(new Tooltip("Click to run"));
        selectBtn.setTooltip(new Tooltip("Select all of the PDF files you want in excel"));
        listViewStrings.setTooltip(new Tooltip("List of all selected files"));
        moveUp.setTooltip(new Tooltip("Move's selected file up the list"));
        moveDown.setTooltip(new Tooltip("Move's selected file down the list"));

        GridPane.setHalignment(okBtn, HPos.RIGHT);

        root.add(lbl, 0, 0);
        root.add(author, 2, 5);
        root.add(field, 1, 0, 3, 1);
        root.add(listViewStrings, 0, 1, 4, 2);
        root.add(okBtn, 3, 3);
        root.add(selectBtn, 2, 3);
        root.add(deleteBtn, 0, 5);
        root.add(moveUp, 0, 3);
        root.add(moveDown, 0, 4);

        stage.setTitle("ICP-MS PerkinElmer PDF to Excel");
        stage.setScene(new Scene(root, 525, 300));
        stage.show();

        savedStage = stage;
    }
    
    private void secondarySelecteFile() {
                FileChooser addFile = new FileChooser();
                addFile.setTitle("Select PDF Files");
                
                System.out.println("Selecting more files to add");
                
                addFile.getExtensionFilters().addAll(new ExtensionFilter("PDF Files", "*.pdf"));
                List<File> selectedFiles = addFile.showOpenMultipleDialog(savedStage); 
                
                for (File selectedFile : selectedFiles) {  // gets all of the path's to selected files and saves them as a string
                    String tempFilePath = selectedFile.getAbsolutePath();
                    System.out.println(tempFilePath);
                     selStrings.add(tempFilePath);
                
                System.out.println("Selected Files: " + selStrings);
                System.out.println("# of Files: " + selStrings.size());
            }
                
    }
    
    private void shortenName() {
        
        System.out.println("Before Clear: " + shortStrings);
        if(shortStrings.isEmpty() == false){
            shortStrings.clear();
        }
        System.out.println("After Clear: " + shortStrings);
        for (int i= 0; i < selStrings.size(); i++) {
                String tempFilePath = selStrings.get(i);
               if(OStype.contains("Mac")){
                
                   String slash = System.getProperty("file.separator");
                   String result[] = tempFilePath.split(slash);
                   
                   System.out.println("File Separator: "+ slash);
                   
                String shortFilePath = result[result.length - 3] + slash + result[result.length - 2] + slash + result[result.length - 1];
                shortStrings.add(shortFilePath);
                System.out.println(shortFilePath);
               }
               if(OStype.contains("Windows")){
                shortStrings.add(selStrings.get(i));
                System.out.println(selStrings);
               }
             
            }
    }

    
    
    private void refreshListView() { //updates list view to show any changes(Move up, move down, delete)
        listViewStrings.getItems().clear();
        
        for (String shortStrings : shortStrings) {
            listViewStrings.getItems().add(shortStrings);
        }
    }

    private void moveUp(int selectedIndex) {
        /*
        get selectedIndex - 1
        get selectedString
        get selectedString -1
        save selectedString to tempstring
        move selectedString -1 to selectedString index
        put tempString at selected string -1 index
        refresh selStrings
        refresh listViewStrings
        */
        int replaceIndex = selectedIndex - 1;
            String selectedString = selStrings.get(selectedIndex);
            String swapString = selStrings.get(replaceIndex);
            String tempStringA = selectedString;
            String tempStringB = swapString;
            selStrings.remove(selectedIndex);
            selStrings.add(selectedIndex, tempStringB);
            selStrings.remove(replaceIndex);
            selStrings.add(replaceIndex, tempStringA);
            
            String selectedShortString = shortStrings.get(selectedIndex);
            String swapShortString = shortStrings.get(replaceIndex);
            String tempShortStringA = selectedShortString;
            String tempShortStringB = swapShortString;
            shortStrings.remove(selectedIndex);
            shortStrings.add(selectedIndex, tempShortStringB);
            shortStrings.remove(replaceIndex);
            shortStrings.add(replaceIndex, tempShortStringA);
            
        refreshListView();
        listViewStrings.getSelectionModel().select(selectedIndex - 1);
    }

    private void moveDown(int selectedIndex) {
        int replaceIndex = selectedIndex + 1;
            String selectedString = selStrings.get(selectedIndex);
            String swapString = selStrings.get(replaceIndex);
            String tempStringA = selectedString;
            String tempStringB = swapString;
            selStrings.remove(selectedIndex);
            selStrings.add(selectedIndex, tempStringB);
            selStrings.remove(replaceIndex);
            selStrings.add(replaceIndex, tempStringA);

            String selectedShortString = shortStrings.get(selectedIndex);
            String swapShortString = shortStrings.get(replaceIndex);
            String tempShortStringA = selectedShortString;
            String tempShortStringB = swapShortString;
            shortStrings.remove(selectedIndex);
            shortStrings.add(selectedIndex, tempShortStringB);
            shortStrings.remove(replaceIndex);
            shortStrings.add(replaceIndex, tempShortStringA);
        
        refreshListView();
        listViewStrings.getSelectionModel().select(selectedIndex + 1);
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        launch(args);
    }

    /**
     * ***************************************************************************
     * <p>
     * END OF GUI                                  *
     *******************************************************************************/

    private static int selectedFileSize = 0;
    private static int n = 1;
    private static String[] alphabet = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S", "T", "U", "V", "W", "X", "Y", "Z", "AA","AB","AC","AD","AE", "AF","AG","AH","AI","AJ","AK","AL","AM","AN","AO","AP","AQ","AR","AS","AT","AU","AV","AW","AX","AY","AZ","BA","BB","BC","BD","BE", "BF","BG","BH","BI","BJ","BK","BL","BM","BN","BO","BP","BQ","BR","BS","BT","BU","BV","BW","BX","BY","BZ", "CA","CB","CC","CD","CE", "CF","CG","CH","CI","CJ","CK","CL","CM","CN","CO","CP","CQ","CR","CS","CT","CU","CV","CW","CX","CY","CZ"};

    private static void excelTemplate() {   //creates my excel template that will be filled with the ICP-MS Ions data
        tmp = new File(dirPath + System.getProperty("file.separator") +"Template_Ions.xls");
        tmpPath = tmp.getAbsolutePath();
        boolean exists = tmp.exists();

        if (exists) {
            String templatePath = tmp.getAbsolutePath();
            convert(templatePath);
        } else {

            System.out.println("***********************");
            System.out.println("Creating Excel Template");
            System.out.println("***********************");

            String templatePath = null;
            try {
                File template = new File(dirPath + System.getProperty("file.separator") +"Template_Ions.xls"); //creates the template file
                template.createNewFile();
                try (FileOutputStream ions = new FileOutputStream(template, false)) {
                    HSSFWorkbook workbook = new HSSFWorkbook();
                    HSSFSheet worksheet = workbook.createSheet("Ions");

                    Font fontBold = workbook.createFont();
                    fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
                    Font fontRed = workbook.createFont();
                    fontRed.setColor(HSSFColor.RED.index);

                    //Grey Cell Style
                    HSSFCellStyle greyStyle = workbook.createCellStyle();

                    greyStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                    greyStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    greyStyle.setAlignment(ALIGN_CENTER);
                    greyStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    greyStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    greyStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                    greyStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);
                    
                    //grey Bold style (for critical Ions)
                    HSSFCellStyle greyStyleBold = workbook.createCellStyle();
                    greyStyleBold.setFont(fontBold);
                    greyStyleBold.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                    greyStyleBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
                    greyStyleBold.setAlignment(ALIGN_CENTER);
                    greyStyleBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                    greyStyleBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
                    greyStyleBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
                    greyStyleBold.setBorderRight(HSSFCellStyle.BORDER_THIN);

                    System.out.println("Populating common fields..... ");

                    System.out.println(selectedFileSize);
            /* rest of this method creates each row, then creates the first cell 
               of the column and fills it with the ions template */
                    String[] nameString = {"name", "lot", "stage", "conc", "Be", "Na", "Mg", "Al", "K", "Ca", "Ti", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Ga", "Zr", "Mo", "Ru", "Cd", "In", "Sn", "Li", "Zn", "Sb", "W", "Pb", "row30", "row31", "critHeader", "critLot", "critConc", "critNa", "critMg", "critAl", "critK", "critCa", "critCr", "critMn", "critFe", "critNi", "critCu", "critTot"};

                    System.out.println("Name sting: ");
                    for (String nameString1 : nameString) {
                        System.out.print(nameString1 + ", ");
                    }

                    System.out.println("Cell String: ");
                    String[] cells = new String[50];//cellA1-A46
                    for (int i = 0; i < 46; i++) {
                        cells[i] = "cellA" + (i + 1);
                        System.out.print(cells[i] + ", ");
                    }

                    System.out.println("Text String: "); //Could add customizable Excels by user inputting template names for String[] text
                    String[] text = {"Name: ", "Lot #: ", "Stage: ", "Analyte: ", "Be", "Na", "Mg", "Al", "K", "Ca", "Ti", "Cr", "Mn", "Fe", "Co", "Ni", "Cu", "Ga", "Zr", "Mo", "Ru", "Cd", "In", "Sn", "Li", "Zn", "Sb", "W", "Pb", "row30", "row31", "10 Critical Ions", "Lot: ", "Conc: ", "Na", "Mg", "Al", "K", "Ca", "Cr", "Mn", "Fe", "Ni", "Cu", "Total: "};
                    for (int i = 0; i < nameString.length; ++i) {
                        System.out.print(text[i] + ", ");
                    }
                    System.out.println("**");

                    System.out.println("Text list is " + text.length + " indexies long.");

                   
                    
                    HSSFRow name = worksheet.createRow(0);   //creates row 1
                    HSSFCell cellA1 = name.createCell(0); // creates cell A1
                    cellA1.setCellValue("Name: ");   //sets value of cell
                    cellA1.setCellStyle(greyStyleBold); //sets cell style(bold for either header or critical ion)

                    HSSFRow lot = worksheet.createRow(1);
                    HSSFCell cellA2 = lot.createCell(0);
                    cellA2.setCellValue("Lot #: ");
                    cellA2.setCellStyle(greyStyleBold);

                    HSSFRow stage = worksheet.createRow(2);
                    HSSFCell cellA3 = stage.createCell(0);
                    cellA3.setCellValue("Stage: ");
                    cellA3.setCellStyle(greyStyleBold);

                    HSSFRow conc = worksheet.createRow(3);
                    HSSFCell cellA4 = conc.createCell(0);
                    cellA4.setCellValue("Analyte");
                    cellA4.setCellStyle(greyStyleBold);

                    HSSFRow Be = worksheet.createRow(4);
                    HSSFCell cellA5 = Be.createCell(0);
                    cellA5.setCellValue("Be");
                    cellA5.setCellStyle(greyStyle);

                    HSSFRow Na = worksheet.createRow(5);
                    HSSFCell cellA6 = Na.createCell(0);
                    cellA6.setCellValue("Na");
                    cellA6.setCellStyle(greyStyleBold);

                    HSSFRow Mg = worksheet.createRow(6);
                    HSSFCell cellA7 = Mg.createCell(0);
                    cellA7.setCellValue("Mg");
                    cellA7.setCellStyle(greyStyleBold);

                    HSSFRow Al = worksheet.createRow(7);
                    HSSFCell cellA8 = Al.createCell(0);
                    cellA8.setCellValue("Al");
                    cellA8.setCellStyle(greyStyleBold);

                    HSSFRow K = worksheet.createRow(8);
                    HSSFCell cellA9 = K.createCell(0);
                    cellA9.setCellValue("K");
                    cellA9.setCellStyle(greyStyleBold);

                    HSSFRow Ca = worksheet.createRow(9);
                    HSSFCell cellA10 = Ca.createCell(0);
                    cellA10.setCellValue("Ca");
                    cellA10.setCellStyle(greyStyleBold);

                    HSSFRow Ti = worksheet.createRow(10);
                    HSSFCell cellA11 = Ti.createCell(0);
                    cellA11.setCellValue("Ti");
                    cellA11.setCellStyle(greyStyle);

                    HSSFRow Cr = worksheet.createRow(11);
                    HSSFCell cellA12 = Cr.createCell(0);
                    cellA12.setCellValue("Cr");
                    cellA12.setCellStyle(greyStyleBold);

                    HSSFRow Mn = worksheet.createRow(12);
                    HSSFCell cellA13 = Mn.createCell(0);
                    cellA13.setCellValue("Mn");
                    cellA13.setCellStyle(greyStyleBold);

                    HSSFRow Fe = worksheet.createRow(13);
                    HSSFCell cellA14 = Fe.createCell(0);
                    cellA14.setCellValue("Fe");
                    cellA14.setCellStyle(greyStyleBold);

                    HSSFRow Co = worksheet.createRow(14);
                    HSSFCell cellA15 = Co.createCell(0);
                    cellA15.setCellValue("Co");
                    cellA15.setCellStyle(greyStyle);

                    HSSFRow Ni = worksheet.createRow(15);
                    HSSFCell cellA16 = Ni.createCell(0);
                    cellA16.setCellValue("Ni");
                    cellA16.setCellStyle(greyStyleBold);

                    HSSFRow Cu = worksheet.createRow(16);
                    HSSFCell cellA17 = Cu.createCell(0);
                    cellA17.setCellValue("Cu");
                    cellA17.setCellStyle(greyStyleBold);

                    HSSFRow Ga = worksheet.createRow(17);
                    HSSFCell cellA18 = Ga.createCell(0);
                    cellA18.setCellValue("Ga");
                    cellA18.setCellStyle(greyStyle);

                    HSSFRow Zr = worksheet.createRow(18);
                    HSSFCell cellA19 = Zr.createCell(0);
                    cellA19.setCellValue("Zr");
                    cellA19.setCellStyle(greyStyle);

                    HSSFRow Mo = worksheet.createRow(19);
                    HSSFCell cellA20 = Mo.createCell(0);
                    cellA20.setCellValue("Mo");
                    cellA20.setCellStyle(greyStyle);

                    HSSFRow Ru = worksheet.createRow(20);
                    HSSFCell cellA21 = Ru.createCell(0);
                    cellA21.setCellValue("Ru");
                    cellA21.setCellStyle(greyStyle);

                    HSSFRow Cd = worksheet.createRow(21);
                    HSSFCell cellA22 = Cd.createCell(0);
                    cellA22.setCellValue("Cd");
                    cellA22.setCellStyle(greyStyle);

                    HSSFRow In = worksheet.createRow(22);
                    HSSFCell cellA23 = In.createCell(0);
                    cellA23.setCellValue("In");
                    cellA23.setCellStyle(greyStyle);

                    HSSFRow Sn = worksheet.createRow(23);
                    HSSFCell cellA24 = Sn.createCell(0);
                    cellA24.setCellValue("Sn");
                    cellA24.setCellStyle(greyStyle);

                    HSSFRow Li = worksheet.createRow(24);
                    HSSFCell cellA25 = Li.createCell(0);
                    cellA25.setCellValue("Li");
                    cellA25.setCellStyle(greyStyle);

                    HSSFRow Zn = worksheet.createRow(25);
                    HSSFCell cellA26 = Zn.createCell(0);
                    cellA26.setCellValue("Zn");
                    cellA26.setCellStyle(greyStyle);

                    HSSFRow Sb = worksheet.createRow(26);
                    HSSFCell cellA27 = Sb.createCell(0);
                    cellA27.setCellValue("Sb");
                    cellA27.setCellStyle(greyStyle);

                    HSSFRow W = worksheet.createRow(27);
                    HSSFCell cellA28 = W.createCell(0);
                    cellA28.setCellValue("W");
                    cellA28.setCellStyle(greyStyle);

                    HSSFRow Pb = worksheet.createRow(28);
                    HSSFCell cellA29 = Pb.createCell(0);
                    cellA29.setCellValue("Pb");
                    cellA29.setCellStyle(greyStyle);

                    HSSFRow tot = worksheet.createRow(29);
                    HSSFCell cellA30 = tot.createCell(0);
                    cellA30.setCellValue("Total: ");
                    cellA30.setCellStyle(greyStyle);

                    HSSFRow row30 = worksheet.createRow(30);
                    HSSFCell cellA31 = row30.createCell(0);

                    HSSFRow row31 = worksheet.createRow(31);
                    HSSFCell cellA32 = row30.createCell(0);

                    HSSFRow critHeader = worksheet.createRow(32);
                    HSSFCell cellA33 = critHeader.createCell(0);
                    cellA33.setCellValue("13 Critical Ions");
                    worksheet.addMergedRegion(new CellRangeAddress(32, 32, 0, size));
                    cellA33.setCellStyle(greyStyle);

                    HSSFRow critLot = worksheet.createRow(33);
                    HSSFCell cellA34 = critLot.createCell(0);
                    cellA34.setCellValue("Lot: ");
                    cellA34.setCellStyle(greyStyle);

                    HSSFRow critConc = worksheet.createRow(34);
                    HSSFCell cellA35 = critConc.createCell(0);
                    cellA35.setCellValue("Analyte");
                    cellA35.setCellStyle(greyStyle);

                    HSSFRow critNa = worksheet.createRow(35);
                    HSSFCell cellA36 = critNa.createCell(0);
                    cellA36.setCellValue("Na");
                    cellA36.setCellStyle(greyStyle);

                    HSSFRow critMg = worksheet.createRow(36);
                    HSSFCell cellA37 = critMg.createCell(0);
                    cellA37.setCellValue("Mg");
                    cellA37.setCellStyle(greyStyle);

                    HSSFRow critAl = worksheet.createRow(37);
                    HSSFCell cellA38 = critAl.createCell(0);
                    cellA38.setCellValue("Al");
                    cellA38.setCellStyle(greyStyle);

                    HSSFRow critK = worksheet.createRow(38);
                    HSSFCell cellA39 = critK.createCell(0);
                    cellA39.setCellValue("K");
                    cellA39.setCellStyle(greyStyle);

                    HSSFRow critCa = worksheet.createRow(39);
                    HSSFCell cellA40 = critCa.createCell(0);
                    cellA40.setCellValue("Ca");
                    cellA40.setCellStyle(greyStyle);

                    HSSFRow critCr = worksheet.createRow(40);
                    HSSFCell cellA41 = critCr.createCell(0);
                    cellA41.setCellValue("Cr");
                    cellA41.setCellStyle(greyStyle);

                    HSSFRow critMn = worksheet.createRow(41);
                    HSSFCell cellA42 = critMn.createCell(0);
                    cellA42.setCellValue("Mn");
                    cellA42.setCellStyle(greyStyle);

                    HSSFRow critFe = worksheet.createRow(42);
                    HSSFCell cellA43 = critFe.createCell(0);
                    cellA43.setCellValue("Fe");
                    cellA43.setCellStyle(greyStyle);

                    HSSFRow critNi = worksheet.createRow(43);
                    HSSFCell cellA44 = critNi.createCell(0);
                    cellA44.setCellValue("Ni");
                    cellA44.setCellStyle(greyStyle);

                    HSSFRow critCu = worksheet.createRow(44);
                    HSSFCell cellA45 = critCu.createCell(0);
                    cellA45.setCellValue("Cu");
                    cellA45.setCellStyle(greyStyle);

                    HSSFRow critSn = worksheet.createRow(45);
                    HSSFCell cellA46 = critSn.createCell(0);
                    cellA46.setCellValue("Sn");
                    cellA46.setCellStyle(greyStyle);
                    
                    HSSFRow critZn = worksheet.createRow(46);
                    HSSFCell cellA47 = critZn.createCell(0);
                    cellA47.setCellValue("Zn");
                    cellA47.setCellStyle(greyStyle);
                    
                    HSSFRow critTi = worksheet.createRow(47);
                    HSSFCell cellA48 = critTi.createCell(0);
                    cellA48.setCellValue("Ti");
                    cellA48.setCellStyle(greyStyle);
                    
                    HSSFRow critTot = worksheet.createRow(48);
                    HSSFCell cellA49 = critTot.createCell(0);
                    cellA49.setCellValue("Total: ");
                    cellA49.setCellStyle(greyStyle); 

                    workbook.write(ions);
                    ions.flush();
                    ions.close();
                    
                    
                            
                } 
                    
               

                templatePath = template.getAbsolutePath();
                convert(templatePath);
            } catch (IOException ex) {
                Logger.getLogger(PDFtoExcel.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    private static void convert(String templatePath) { //runs program on order that the PDF's are listed in listview
        //this allows excel to be filled in the desired order
        for (String selString : selStrings) {
            System.out.println("LOOK AT ME " + selString);
            pdfToTxt(selString, templatePath);
        }
    }

    private static void pdfToTxt(String fPath, String templatePath) {
        selectedFileSize = size;

        try {
            File input = new File(fPath);  // The PDF file from where you would like to extract
            File output = new File(dirPath + System.getProperty("file.separator") +"output.txt");
            
            
            
            PDDocument pd = PDDocument.load(input);
            System.out.println(pd.getNumberOfPages());
            System.out.println(pd.isEncrypted());
            pd.save(dirPath + System.getProperty("file.separator") +"IonsCopy.pdf"); // Creates a copy of pdf
            
            
            PDFTextStripper stripper = new PDFTextStripper();

            BufferedWriter wr = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output)));
            stripper.writeText(pd, wr); //strips all text from PDF document and wrights it to the Buffered Writer
            if (pd != null) {
                pd.close();
            }

            wr.close();
            
            txtToCsv(templatePath);
        } catch (Exception e) {
        }
    }

    private static void txtToCsv(String templatePath) throws IOException {
        File file = new File(dirPath + System.getProperty("file.separator") +"output.txt"); //grabs text file from before
        Scanner scan = new Scanner(file);
        csvFile = new File(dirPath + System.getProperty("file.separator") +"CSV.csv");  //creates new CSV file
        file.createNewFile();
        csvFilePath = csvFile.getAbsolutePath();
        FileWriter writer = new FileWriter(csvFile);

        while (scan.hasNext()) {

            String csv = scan.nextLine().replace(" ", ","); //scans through text file, replaces all spaces with commas

            System.out.println(csv);
            System.out.println("Length: " + csv.length());
            writer.append(csv);
            writer.append("\n");
            writer.flush();
            
        }
        scan.close();
        writer.close();
        file.delete();
        getData(templatePath);
    }

    private static void getData(String templatePath) throws IOException {
        System.out.println("******************************");
        String stage = null;    //initializing all strings needed below
        String a = null;
        String Be = null;
        String Na = null;
        String Mg = null;
        String Al = null;
        String K = null;
        String Ca = null;
        String Ti = null;
        String Cr = null;
        String Mn = null;
        String Fe = null;
        String Co = null;
        String Ni = null;
        String Cu = null;
        String Ga = null;
        String Zr = null;
        String Mo = null;
        String Ru = null;
        String Cd = null;
        String In = null;
        String Sn = null;
        String Li = null;
        String Zn = null;
        String Sb = null;
        String W = null;
        String Pb = null;
        String material = null;
        String lotNum = null;
        String analyte = null;
        String lot = null;

        List<String> list = new ArrayList<>();
        List<String> Ion = new ArrayList<>();

        FileResource csv = new FileResource(dirPath + System.getProperty("file.separator") +"CSV.csv"); //grabs previously created CSV file
        CSVParser parser = csv.getCSVParser(false);
        for (CSVRecord record : parser) { //Scans CSV
            a = record.get(0); //scans first column of CSV
            
            if (a.contains("Material:")) { //if desired word is in first column of CSV
                System.out.println("Perkin Elmer File");
                System.out.println(a + " " + record.get(1));
                material = record.get(1); //get the item in the next column over on the same row
                Ion.add(record.get(0));  //get the desired word
                list.add(material);   // adds item to list
            }
            
            if (a.contains("Lot")) {
                System.out.println(a + "" + record.get(2));
                 
                lot = record.get(2);
                    if(lot.isEmpty()){
                        lot = ("Lot Number");
                    }; 
                
                list.add(lot);
                Ion.add(record.get(0));
                lotNum = record.get(2);
            }
            if (a.contains("Stage")) {
               
                System.out.println(a + "" + record.get(2));
                System.out.println("Size is: " + record.size());
                
                String tmpstage = null; 
                
                System.out.println("Tmpstage: " + tmpstage);
                
                stage = record.get(1) + " " + record.get(2);
                list.add(stage);
                Ion.add(record.get(0));
                System.out.println(list);
            }
            if (a.contains("Analyte")) {
                System.out.println(a + "" + record.get(6));
                analyte = record.get(6);
                list.add(analyte);
                Ion.add(record.get(0));
            }

            if (a.contains("Be") && a.length() <= 3) { //if the Ion is what I'm looking for 
                Be = record.get(3);                  //and its only 2 chars long
                if(Be.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Be);                        //add the resulting conc to list
                Ion.add(record.get(0)); 
                }//add the ion name to Ion
            }

            if (a.contains("Na") && a.length() <= 3) {
                Na = record.get(3);
                if(Na.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Na);
                Ion.add(record.get(0));
                }
                
            }
            if (a.contains("Mg") && a.length() <= 3) {
                Mg = record.get(3);
                if(Mg.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Mg);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Al") && a.length() <= 3) {
                Al = record.get(3);
                if(Al.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Al);
                Ion.add(record.get(0));
                }
                
            }
            if (a.contains("K") && a.length() <= 3) {
                K = record.get(3);
                if(K.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(K);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Ca") && a.length() <= 3) {
                Ca = record.get(3);
                if(Ca.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Ca);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Ti") && a.length() <= 3) {
                Ti = record.get(3);
                if(Ti.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Ti);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Cr") && a.length() <= 3) {
                Cr = record.get(3);
                if(Cr.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Cr);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Mn") && a.length() <= 3) {
                Mn = record.get(3);
                if(Mn.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Mn);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Fe") && a.length() <= 3) {
                Fe = record.get(3);
                if(Fe.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Fe);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Co") && a.length() <= 3) {
                Co = record.get(3);
                if(Co.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Co);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Ni") && a.length() <= 3) {
                Ni = record.get(3);
                if(Ni.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Ni);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Cu") && a.length() <= 3) {
                Cu = record.get(3);
                if(Cu.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Cu);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Ga") && a.length() <= 3) {
                Ga = record.get(3);
                if(Ga.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Ga);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Zr") && a.length() <= 3) {
                Zr = record.get(3);
                if(Zr.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Zr);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Mo") && a.length() <= 3) {
                Mo = record.get(3);
                if(Mo.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Mo);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Ru") && a.length() <= 3) {
                Ru = record.get(3);
                if(Ru.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Ru);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Cd") && a.length() <= 3) {
                Cd = record.get(3);
                if(Cd.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Cd);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("In") && a.length() <= 3) {
                In = record.get(3);
                if(In.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(In);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Sn") && a.length() <= 3) {
                Sn = record.get(3);
                if(Sn.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Sn);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Li") && a.length() <= 3) {
                Li = record.get(3);
                if(Li.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Li);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Zn") && a.length() <= 3) {
                Zn = record.get(3);
                if(Zn.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Zn);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Sb") && a.length() <= 3) {
                Sb = record.get(3);
                if(Sb.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Sb);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("W") && a.length() <= 3) {
                W = record.get(3);
                if(W.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(W);
                Ion.add(record.get(0));
                }
            }
            if (a.contains("Pb") && a.length() <= 3) {
                Pb = record.get(3);
                if(Pb.contains("S")){
                   list.add("Saturated");
                   Ion.add("Saturated");
              } 
                else{
                list.add(Pb);
                Ion.add(record.get(0));
                }
            }
        }
        parser.close();

        addToExcel(list, Ion, material, lotNum, templatePath);
    }

    //addToExcel method takes the parsed data and fills it into the template sheet created earlier
    private static void addToExcel(List list, List Ion, String material, String lotNum, String templatePath) throws IOException {
        System.out.println("***************************************************");
        System.out.println("Starting AddToExcel");
        System.out.println("***************************************************");
        String[] listString = (String[]) list.toArray(new String[0]); //adds all previously grabbed names to list for parsing
        String[] ionString = (String[]) Ion.toArray(new String[0]);
        System.out.println("List: " + list);
        System.out.println("Ions: " + Ion);

        String nameLot = "Name";
        
        System.out.println("Test1");
        
        FileInputStream template = new FileInputStream(new File(templatePath)); //gets template created earlier

        HSSFWorkbook workbook = new HSSFWorkbook(template);
        HSSFSheet worksheet = workbook.getSheetAt(0);
        //Cell and Font Tyles
        Font fontBold = workbook.createFont();
        fontBold.setBoldweight(Font.BOLDWEIGHT_BOLD);
        Font fontRed = workbook.createFont();
        fontRed.setColor(HSSFColor.RED.index);
        Font fontBoldRed = workbook.createFont();
        fontBoldRed.setColor(HSSFColor.RED.index);
        fontBoldRed.setBoldweight(Font.BOLDWEIGHT_BOLD);

        //Grey Cell Style
        HSSFCellStyle greyStyle = workbook.createCellStyle();
        greyStyle.setFont(fontBold);
        greyStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        greyStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        greyStyle.setAlignment(ALIGN_CENTER);
        greyStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        greyStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        greyStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        greyStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);


        //Green Bolded cell  style
        HSSFCellStyle greenStyleBold = workbook.createCellStyle();
        greenStyleBold.setFont(fontBold);
        greenStyleBold.setFillForegroundColor(IndexedColors.LIME.getIndex());
        greenStyleBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
        greenStyleBold.setAlignment(ALIGN_CENTER);
        greenStyleBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        greenStyleBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
        greenStyleBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        greenStyleBold.setBorderRight(HSSFCellStyle.BORDER_THIN);


        //Green cell style
        HSSFCellStyle greenStyle = workbook.createCellStyle();

        greenStyle.setFillForegroundColor(IndexedColors.LIME.getIndex());
        greenStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        greenStyle.setAlignment(ALIGN_CENTER);
        greenStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        greenStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        greenStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        greenStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);

        //Blue Bold cell style
        HSSFCellStyle blueStyleBold = workbook.createCellStyle();
        blueStyleBold.setFont(fontBold);
        blueStyleBold.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        blueStyleBold.setFillPattern(CellStyle.SOLID_FOREGROUND);
        blueStyleBold.setAlignment(ALIGN_CENTER);
        blueStyleBold.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        blueStyleBold.setBorderTop(HSSFCellStyle.BORDER_THIN);
        blueStyleBold.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        blueStyleBold.setBorderRight(HSSFCellStyle.BORDER_THIN);

        //Blue cell style
        HSSFCellStyle blueStyle = workbook.createCellStyle();
        blueStyle.setFillForegroundColor(IndexedColors.AQUA.getIndex());
        blueStyle.setFillPattern(CellStyle.SOLID_FOREGROUND);
        blueStyle.setAlignment(ALIGN_CENTER);
        blueStyle.setBorderBottom(HSSFCellStyle.BORDER_THIN);
        blueStyle.setBorderTop(HSSFCellStyle.BORDER_THIN);
        blueStyle.setBorderLeft(HSSFCellStyle.BORDER_THIN);
        blueStyle.setBorderRight(HSSFCellStyle.BORDER_THIN);

        System.out.println("n = " + n);

        System.out.println("Files to work with: ");
        System.out.println(selectedFileSize + ", on number " + n + " ");

        int totalRowNum = worksheet.getPhysicalNumberOfRows() + 3;
        System.out.println("Last row at: " + totalRowNum);
            
            /*gets all rows and cells from template excel sheet
            Can't loop over these because the name of each row and cell 
            variable change each time                                  */

        HSSFRow name = worksheet.getRow(0);
        HSSFCell cellA1 = name.getCell(0);


        System.out.println("661");
        HSSFRow lot = worksheet.getRow(1);
        HSSFCell cellA2 = name.getCell(0);

        System.out.println("665");
        HSSFRow stage = worksheet.getRow(2);
        HSSFCell cellA3 = name.getCell(0);

        HSSFRow conc = worksheet.getRow(3);
        HSSFCell cellA4 = conc.getCell(0);

        HSSFRow Be = worksheet.getRow(4);
        HSSFCell cellA5 = Be.getCell(0);

        HSSFRow Na = worksheet.getRow(5);
        HSSFCell cellA6 = Na.getCell(0);

        HSSFRow Mg = worksheet.getRow(6);
        HSSFCell cellA7 = Mg.getCell(0);

        HSSFRow Al = worksheet.getRow(7);
        HSSFCell cellA8 = Al.getCell(0);

        HSSFRow K = worksheet.getRow(8);
        HSSFCell cellA9 = K.getCell(0);

        HSSFRow Ca = worksheet.getRow(9);
        HSSFCell cellA10 = Ca.getCell(0);

        HSSFRow Ti = worksheet.getRow(10);
        HSSFCell cellA11 = Ti.getCell(0);

        HSSFRow Cr = worksheet.getRow(11);
        HSSFCell cellA12 = Cr.getCell(0);

        HSSFRow Mn = worksheet.getRow(12);
        HSSFCell cellA13 = Mn.getCell(0);

        HSSFRow Fe = worksheet.getRow(13);
        HSSFCell cellA14 = Fe.getCell(0);

        HSSFRow Co = worksheet.getRow(14);
        HSSFCell cellA15 = Co.getCell(0);

        HSSFRow Ni = worksheet.getRow(15);
        HSSFCell cellA16 = Ni.getCell(0);

        HSSFRow Cu = worksheet.getRow(16);
        HSSFCell cellA17 = Cu.getCell(0);

        HSSFRow Ga = worksheet.getRow(17);
        HSSFCell cellA18 = Ga.getCell(0);

        HSSFRow Zr = worksheet.getRow(18);
        HSSFCell cellA19 = Zr.getCell(0);

        HSSFRow Mo = worksheet.getRow(19);
        HSSFCell cellA20 = Mo.getCell(0);

        HSSFRow Ru = worksheet.getRow(20);
        HSSFCell cellA21 = Ru.getCell(0);

        HSSFRow Cd = worksheet.getRow(21);
        HSSFCell cellA22 = Cd.getCell(0);

        HSSFRow In = worksheet.getRow(22);
        HSSFCell cellA23 = In.getCell(0);

        HSSFRow Sn = worksheet.getRow(23);
        HSSFCell cellA24 = Sn.getCell(0);

        HSSFRow Li = worksheet.getRow(24);
        HSSFCell cellA25 = Li.getCell(0);

        HSSFRow Zn = worksheet.getRow(25);
        HSSFCell cellA26 = Zn.getCell(0);

        HSSFRow Sb = worksheet.getRow(26);
        HSSFCell cellA27 = Sb.getCell(0);

        HSSFRow W = worksheet.getRow(27);
        HSSFCell cellA28 = W.getCell(0);

        HSSFRow Pb = worksheet.getRow(28);
        HSSFCell cellA29 = Pb.getCell(0);

        HSSFRow tot = worksheet.getRow(29);
        HSSFCell cellA30 = tot.getCell(0);

        HSSFRow row31 = worksheet.getRow(30);
        HSSFCell cellA31 = row31.getCell(0);

        HSSFRow row32 = worksheet.getRow(31);
        HSSFCell cellA32 = row32.getCell(0);

        System.out.println("Now starting Crit Ions");

        HSSFRow critHeader = worksheet.getRow(32);
        HSSFCell cellA33 = critHeader.getCell(0);
        System.out.println("line 762");
        HSSFRow critLot = worksheet.getRow(33);
        HSSFCell cellA34 = critLot.getCell(0);

        HSSFRow critConc = worksheet.getRow(34);
        HSSFCell cellA35 = critConc.getCell(0);

        HSSFRow critNa = worksheet.getRow(35);
        HSSFCell cellA36 = critNa.getCell(0);

        HSSFRow critMg = worksheet.getRow(36);
        HSSFCell cellA37 = critMg.getCell(0);

        HSSFRow critAl = worksheet.getRow(37);
        HSSFCell cellA38 = critAl.getCell(0);

        HSSFRow critK = worksheet.getRow(38);
        HSSFCell cellA39 = critK.getCell(0);

        HSSFRow critCa = worksheet.getRow(39);
        HSSFCell cellA40 = critCa.getCell(0);

        HSSFRow critCr = worksheet.getRow(40);
        HSSFCell cellA41 = critCr.getCell(0);

        HSSFRow critMn = worksheet.getRow(41);
        HSSFCell cellA42 = critMn.getCell(0);

        HSSFRow critFe = worksheet.getRow(42);
        HSSFCell cellA43 = critFe.getCell(0);

        HSSFRow critNi = worksheet.getRow(43);
        HSSFCell cellA44 = critNi.getCell(0);

        HSSFRow critCu = worksheet.getRow(44);
        HSSFCell cellA45 = critCu.getCell(0);
        
        HSSFRow critSn = worksheet.getRow(45);
        HSSFCell cellA46 = critSn.getCell(0);
        
        HSSFRow critZn = worksheet.getRow(46);
        HSSFCell cellA47 = critZn.getCell(0);
        
        HSSFRow critTi = worksheet.getRow(47);
        HSSFCell cellA48 = critTi.getCell(0);
        
        HSSFRow critTot = worksheet.getRow(48);
        System.out.println("line 817");
        HSSFCell cellA49 = critTot.getCell(0);
        System.out.println("Line 819");

        System.out.println("Gathered all rows and cells");

        System.out.println("Now filling data from the pdf");
        System.out.println("Size: " + size);


        //fills data into excel template sheet
        for (int r = 0; r < size; r++) { // check each row
            if (n - 1 == size) {
                break;
            }
            Row rw = worksheet.getRow(r); //gets each row

            System.out.println("Row Number: " + (r + 1));
            if (rw == null) {
                System.out.println("Row ERROR");
                continue;
            }
            System.out.println("No Row errors: ");
            for (int x = 0; x < totalRowNum; x++) { //check each cell

                Cell c = rw.getCell(x); //gets each cell

                if (c == null) {    //if cell is null, make it Blank
                    c = rw.getCell(x, Row.CREATE_NULL_AS_BLANK);//eliminates null pointers
                    System.out.println("Converting Null cells to Blank");
                }

                System.out.println("No Cell Errors");

                if (c.getCellType() == Cell.CELL_TYPE_BLANK) { //if cell is blank
                    System.out.println("No more Null Cells");
                    //fill blank cell
                    int i = 0;
                    System.out.println("Populating spreadsheet... " + n);

                    // index from 0,0... cell A1 is cell(0,0)
                    HSSFCell cellB1 = name.createCell(n);
                    cellB1.setCellValue(listString[i]);
                    cellB1.setCellStyle(greyStyle);
                    i++;

                    HSSFCell cellB2 = lot.createCell(n);
                    cellB2.setCellValue(listString[i]);
                    cellB2.setCellStyle(greyStyle);
                    i++;

                    HSSFCell cellB3 = stage.createCell(n);
                    cellB3.setCellValue(listString[i]);
                    cellB3.setCellStyle(greyStyle);
                    i++;

                    HSSFCell cellB4 = conc.createCell(n);
                    cellB4.setCellValue("Conc. (ppb)");
                    cellB4.setCellStyle(greyStyle);
                    i++;

                    //this is repeated for 700 lines of code... cant loop again
                    //because cell variable names change each time.

                    HSSFCell cellB5 = Be.createCell(n);
                    //If value is positive, use value, else use 0
                    double zero = 0.00;
                    
                     if(listString[i].contains("S")){
                        cellB5.setCellValue("Saturated");
                    }
                    else{
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB5.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB5.setCellValue(zero); //sets negative numbers to 0,
                        //You cant have a negative concentration of Ions
                     }
                    }
                    //Alternate colors for easier viewing
                    if (n % 2 == 0) {
                        cellB5.setCellStyle(greenStyle);
                    } else {
                        cellB5.setCellStyle(blueStyle);
                    }
                    i++;


                    HSSFCell cellB6 = Na.createCell(n);
                    if(listString[i].contains("S")){
                        cellB6.setCellValue("Saturated");
                    }
                    else{
                    cellB6.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB6.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB6.setCellValue(zero);
                     }
                    }
                    if (n % 2 == 0) {
                        cellB6.setCellStyle(greenStyleBold);
                    } else {
                        cellB6.setCellStyle(blueStyleBold);
                     }
                    
                    i++;


                    HSSFCell cellB7 = Mg.createCell(n);
                    if(listString[i].contains("S")){
                        cellB7.setCellValue("Saturated");
                    }
                    else{
                    cellB7.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB7.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB7.setCellValue(zero);
                     }
                    }
                    if (n % 2 == 0) {
                        cellB7.setCellStyle(greenStyleBold);
                    } else {
                        cellB7.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB8 = Al.createCell(n);
                    if(listString[i].contains("S")){
                        cellB8.setCellValue("Saturated");
                    }
                    else{
                    cellB8.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB8.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB8.setCellValue(zero);
                    }
                    }
                    if (n % 2 == 0) {
                        cellB8.setCellStyle(greenStyleBold);
                    } else {
                        cellB8.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB9 = K.createCell(n);
                    if(listString[i].contains("S")){
                        cellB9.setCellValue("Saturated");
                    }
                    else{
                    cellB9.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB9.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB9.setCellValue(zero);
                    }
                    }
                    if (n % 2 == 0) {
                        cellB9.setCellStyle(greenStyleBold);
                    } else {
                        cellB9.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB10 = Ca.createCell(n);
                    if(listString[i].contains("S")){
                        cellB10.setCellValue("Saturated");
                    }
                    else{
                    cellB10.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB10.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB10.setCellValue(zero);
                     }
                    }
                    if (n % 2 == 0) {
                        cellB10.setCellStyle(greenStyleBold);
                    } else {
                        cellB10.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB11 = Ti.createCell(n);
                    if(listString[i].contains("S")){
                        cellB11.setCellValue("Saturated");
                    }
                    else{
                    cellB11.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB11.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB11.setCellValue(zero);
                    }
                    }
                    
                    if (n % 2 == 0) {
                        cellB11.setCellStyle(greenStyle);
                    } else {
                        cellB11.setCellStyle(blueStyle);
                    }
                    i++;


                    HSSFCell cellB12 = Cr.createCell(n);
                    if(listString[i].contains("S")){
                        cellB12.setCellValue("Saturated");
                    }
                    else{
                    cellB12.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB12.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB12.setCellValue(zero);
                     }
                    }
                    if (n % 2 == 0) {
                        cellB12.setCellStyle(greenStyleBold);
                    } else {
                        cellB12.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB13 = Mn.createCell(n);
                    if(listString[i].contains("S")){
                        cellB13.setCellValue("Saturated");
                    }
                    else{
                    cellB13.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB13.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB13.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB13.setCellStyle(greenStyleBold);
                    } else {
                        cellB13.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB14 = Fe.createCell(n);
                    if(listString[i].contains("S")){
                        cellB14.setCellValue("Saturated");
                    }
                    else{
                    cellB14.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB14.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB14.setCellValue(zero);
                     }
                    }
                    if (n % 2 == 0) {
                        cellB14.setCellStyle(greenStyleBold);
                    } else {
                        cellB14.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB15 = Co.createCell(n);
                    if(listString[i].contains("S")){
                        cellB15.setCellValue("Saturated");
                    }
                    else{
                    cellB15.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB15.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB15.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB15.setCellStyle(greenStyle);
                    } else {
                        cellB15.setCellStyle(blueStyle);
                    }
                    i++;


                    HSSFCell cellB16 = Ni.createCell(n);
                    if(listString[i].contains("S")){
                        cellB16.setCellValue("Saturated");
                    }
                    else{
                    cellB16.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB16.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB16.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB16.setCellStyle(greenStyleBold);
                    } else {
                        cellB16.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB17 = Cu.createCell(n);
                    if(listString[i].contains("S")){
                        cellB17.setCellValue("Saturated");
                    }
                    else{
                    cellB17.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB17.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB17.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB17.setCellStyle(greenStyleBold);
                    } else {
                        cellB17.setCellStyle(blueStyleBold);
                    }
                    i++;


                    HSSFCell cellB18 = Ga.createCell(n);
                    if(listString[i].contains("S")){
                        cellB18.setCellValue("Saturated");
                    }
                    else{
                    cellB18.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB18.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB18.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB18.setCellStyle(greenStyle);
                    } else {
                        cellB18.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB19 = Zr.createCell(n);
                    if(listString[i].contains("S")){
                        cellB9.setCellValue("Saturated");
                    }
                    else{
                    cellB19.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB19.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB19.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB19.setCellStyle(greenStyle);
                    } else {
                        cellB19.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB20 = Mo.createCell(n);
                    if(listString[i].contains("S")){
                        cellB20.setCellValue("Saturated");
                    }
                    else{
                    cellB20.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB20.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB20.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB20.setCellStyle(greenStyle);
                    } else {
                        cellB20.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB21 = Ru.createCell(n);
                    if(listString[i].contains("S")){
                        cellB21.setCellValue("Saturated");
                    }
                    else{
                    cellB21.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB21.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB21.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB21.setCellStyle(greenStyle);
                    } else {
                        cellB21.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB22 = Cd.createCell(n);
                    if(listString[i].contains("S")){
                        cellB22.setCellValue("Saturated");
                    }
                    else{
                    cellB22.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB22.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB22.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB22.setCellStyle(greenStyle);
                    } else {
                        cellB22.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB23 = In.createCell(n);
                    if(listString[i].contains("S")){
                        cellB23.setCellValue("Saturated");
                    }
                    else{
                    cellB23.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB23.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB23.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB23.setCellStyle(greenStyle);
                    } else {
                        cellB23.setCellStyle(blueStyle);
                    }
                    i++;

                    HSSFCell cellB24 = Sn.createCell(n);
                    if(listString[i].contains("S")){
                        cellB24.setCellValue("Saturated");
                    }
                    else{
                    cellB24.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB24.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB24.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB24.setCellStyle(greenStyle);
                    } else {
                        cellB24.setCellStyle(blueStyle);
                    }

                    i++;

                    HSSFCell cellB25 = Li.createCell(n);
                    if(listString[i].contains("S")){
                        cellB25.setCellValue("Saturated");
                    }
                    else{
                    cellB25.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB25.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB25.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB25.setCellStyle(greenStyle);
                    } else {
                        cellB25.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB26 = Zn.createCell(n);
                    if(listString[i].contains("S")){
                        cellB26.setCellValue("Saturated");
                    }
                    else{
                    cellB26.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB26.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB26.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB26.setCellStyle(greenStyle);
                    } else {
                        cellB26.setCellStyle(blueStyle);
                    }


                    i++;

                    HSSFCell cellB27 = Sb.createCell(n);
                    if(listString[i].contains("S")){
                        cellB27.setCellValue("Saturated");
                    }
                    else{
                    cellB27.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB27.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB27.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB27.setCellStyle(greenStyle);
                    } else {
                        cellB27.setCellStyle(blueStyle);
                    }
                    i++;

                    HSSFCell cellB28 = W.createCell(n);
                    if(listString[i].contains("S")){
                        cellB28.setCellValue("Saturated");
                    }
                    else{
                    cellB28.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB28.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB28.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB28.setCellStyle(greenStyle);
                    } else {
                        cellB28.setCellStyle(blueStyle);
                    }

                    i++;

                    HSSFCell cellB29 = Pb.createCell(n);
                    if(listString[i].contains("S")){
                        cellB29.setCellValue("Saturated");
                    }
                    else{
                    cellB29.setCellValue(Double.parseDouble(listString[i]));
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB29.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB29.setCellValue(zero);
                     }
                    }
                    
                    if (n % 2 == 0) {
                        cellB29.setCellStyle(greenStyle);
                    } else {
                        cellB29.setCellStyle(blueStyle);
                    }

                    HSSFCell cellB30 = tot.createCell(n);
                    cellB30.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String sum = "SUM(" + alphabet[n] + "5:" + alphabet[n] + "29)";
                    cellB30.setCellFormula(sum);
                    if (Double.parseDouble(listString[i]) > 0) {
                        cellB30.setCellValue(Double.parseDouble(listString[i]));
                    } else {
                        cellB30.setCellValue(zero);
                    }

                    if (n % 2 == 0) {
                        cellB30.setCellStyle(greenStyleBold);
                    } else {
                        cellB30.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB34 = critLot.createCell(n);
                    cellB34.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critLotVal = (alphabet[n] + "2");
                    cellB34.setCellFormula(critLotVal);
                    cellB34.setCellStyle(greyStyle);

                    HSSFCell cellB35 = critConc.createCell(n);
                    cellB35.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critConcVal = (alphabet[n] + "4");
                    cellB35.setCellFormula(critConcVal);
                    cellB35.setCellStyle(greyStyle);

                    HSSFCell cellB36 = critNa.createCell(n);
                    cellB36.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critNaVal = (alphabet[n] + "6");
                    cellB36.setCellFormula(critNaVal);
                    if (n % 2 == 0) {
                        cellB36.setCellStyle(greenStyleBold);
                    } else {
                        cellB36.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB37 = critMg.createCell(n);
                    cellB37.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critMgVal = (alphabet[n] + "7");
                    cellB37.setCellFormula(critMgVal);
                    if (n % 2 == 0) {
                        cellB37.setCellStyle(greenStyleBold);
                    } else {
                        cellB37.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB38 = critAl.createCell(n);
                    cellB38.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critAlVal = (alphabet[n] + "8");
                    cellB38.setCellFormula(critAlVal);
                    if (n % 2 == 0) {
                        cellB38.setCellStyle(greenStyleBold);
                    } else {
                        cellB38.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB39 = critK.createCell(n);
                    cellB39.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critKVal = (alphabet[n] + "9");
                    cellB39.setCellFormula(critKVal);
                    if (n % 2 == 0) {
                        cellB39.setCellStyle(greenStyleBold);
                    } else {
                        cellB39.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB40 = critCa.createCell(n);
                    cellB40.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critCaVal = (alphabet[n] + "10");
                    cellB40.setCellFormula(critCaVal);
                    if (n % 2 == 0) {
                        cellB40.setCellStyle(greenStyleBold);
                    } else {
                        cellB40.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB41 = critCr.createCell(n);
                    cellB41.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critCrVal = (alphabet[n] + "12");
                    cellB41.setCellFormula(critCrVal);
                    if (n % 2 == 0) {
                        cellB41.setCellStyle(greenStyleBold);
                    } else {
                        cellB41.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB42 = critMn.createCell(n);
                    cellB42.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critMnVal = (alphabet[n] + "13");
                    cellB42.setCellFormula(critMnVal);
                    if (n % 2 == 0) {
                        cellB42.setCellStyle(greenStyleBold);
                    } else {
                        cellB42.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB43 = critFe.createCell(n);
                    cellB43.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critFeVal = (alphabet[n] + "14");
                    cellB43.setCellFormula(critFeVal);
                    if (n % 2 == 0) {
                        cellB43.setCellStyle(greenStyleBold);
                    } else {
                        cellB43.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB44 = critNi.createCell(n);
                    cellB44.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critNiVal = (alphabet[n] + "16");
                    cellB44.setCellFormula(critNiVal);
                    if (n % 2 == 0) {
                        cellB44.setCellStyle(greenStyleBold);
                    } else {
                        cellB44.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB45 = critCu.createCell(n);
                    cellB45.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critCuVal = (alphabet[n] + "17");
                    cellB45.setCellFormula(critCuVal);
                    if (n % 2 == 0) {
                        cellB45.setCellStyle(greenStyleBold);
                    } else {
                        cellB45.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB46 = critSn.createCell(n);
                    cellB46.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critSnVal = (alphabet[n] + "24");
                    cellB46.setCellFormula(critSnVal);
                    if (n % 2 == 0) {
                        cellB46.setCellStyle(greenStyleBold);
                    } else {
                        cellB46.setCellStyle(blueStyleBold);
                    }
                    
                    HSSFCell cellB47 = critZn.createCell(n);
                    cellB47.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critZnVal = (alphabet[n] + "26");
                    cellB47.setCellFormula(critZnVal);
                    if (n % 2 == 0) {
                        cellB47.setCellStyle(greenStyleBold);
                    } else {
                        cellB47.setCellStyle(blueStyleBold);
                    }
                    
                    HSSFCell cellB48 = critTi.createCell(n);
                    cellB48.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critTiVal = (alphabet[n] + "11");
                    cellB48.setCellFormula(critTiVal);
                    if (n % 2 == 0) {
                        cellB48.setCellStyle(greenStyleBold);
                    } else {
                        cellB48.setCellStyle(blueStyleBold);
                    }

                    HSSFCell cellB49 = critTot.createCell(n);
                    cellB49.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                    String critTotVal = ("SUM(" + alphabet[n] + "36:" + alphabet[n] + "48)");
                    cellB49.setCellFormula(critTotVal);
                    if (n % 2 == 0) {
                        cellB49.setCellStyle(greenStyleBold);
                    } else {
                        cellB49.setCellStyle(blueStyleBold);
                    }

                }

            }

        }
        FileOutputStream fileOut = new FileOutputStream(templatePath); //saves file
        workbook.write(fileOut);
        fileOut.flush();
        template.close();
        fileOut.close();
        

        ++n;

        // if empty, use default name, else user specified name
        String destinationPath = isEmpty ? material + "_" + nameLot + "_Ions.xls" : outputName + ".xls";

        System.out.println("*************Template filled*************");
        System.out.println("Now renaming file for you");

        fileOut.close();
        copy(templatePath, destinationPath);
    }
    
   

    private static void copy(String sourcePath, String destinationPath) throws IOException {
        FileOutputStream fileOut = new FileOutputStream(destinationPath);
        Files.copy(Paths.get(sourcePath), fileOut);  //saves to unique output file
        System.out.println("Your spreadsheet is located at: " + destinationPath);
        System.out.println("****************COMPLETE****************");
        
        if(n > size){
            
            open(destinationPath);
            
        }

        clean();
    }
    
    private static void open(String destinationPath) throws IOException{
        Desktop.getDesktop().open(new File(destinationPath));
        if(OStype.contains("Windows")){ 
            String tempPath =tmp.toString();
                System.out.println("tmp " + tempPath);
            tempPath = csvFile.toString();
                System.out.println("csv " + tempPath);
        
        }//opens completed file
    }
    
    private static void clean() {
      
        Runtime.getRuntime().addShutdownHook(new Thread(() -> {
            //deletes tmp files.
            try{
            File[] files = dir.listFiles();
           
            if(files!=null){
                for(File f: files){
                    f.setWritable(true);
                    if(f.delete() == true){
                    System.out.println(f + " was deleted");
                    }
                    else {
                        System.out.println("Cannot delete");
                    }
                }
            }
            for (File f: files){
                System.out.println(f);
            }
            if(dir.isDirectory()){
        dir.delete();
        System.out.println("Directory is deleted : " + dir.getAbsolutePath());
            }
            }catch(Exception e){
                        e.printStackTrace();
                        }
        }));
    }

    
}
