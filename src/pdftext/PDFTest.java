
/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package pdftext;

import edu.duke.FileResource;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.OutputStreamWriter;
import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Scanner;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.text.PDFTextStripper;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import java.nio.file.Files;
import java.nio.file.Paths;
import org.apache.poi.hssf.util.CellRangeAddress;
import org.apache.poi.ss.util.CellUtil;

/**
 *
 * @au\
 */
public class PDFTest  {

public static int selectedFileSize = 0;
public static int n = 1;
public static String [] alphabet = {"A", "B", "C","D","E","F","G","H","I","J","K","L","M","N","O","P","Q","R","S","T","U","V","W","X","Y","Z"};

    
 public static void main(String[] args){
      // pdfTotxt();  run from gui
     }

    public static void pdfTotxt(String fPath, int size, String templatePath) {
        selectedFileSize = size;
        PDDocument pd;
        BufferedWriter wr;
        try {
            
            File input = new File (fPath);  // The PDF file from where you would like to extract
            
            File output = new File("C:\\PDFTester\\output.txt");// The text file where you are going to store the extracted data
            
            pd = PDDocument.load(input);
            System.out.println(pd.getNumberOfPages());
            System.out.println(pd.isEncrypted());
            pd.save("IonsCopy.pdf"); // Creates a copy called "CopyOfInvoice.pdf"
            PDFTextStripper stripper = new PDFTextStripper();
            
            
            
            wr = new BufferedWriter(new OutputStreamWriter(new FileOutputStream(output)));
            stripper.writeText(pd, wr);
            if (pd != null) {
                pd.close();
            }
            
            wr.close();
            txtToCsv(size, templatePath);
        } catch (Exception e){
        }
    }

    
 
 public static void txtToCsv(int size, String templatePath) throws FileNotFoundException, IOException{
     FileWriter writer = null;
      
        File file = new File("C:\\PDFTester\\output.txt");
        Scanner scan = new Scanner(file);
        File file2 = new File("C:\\PDFTester\\CSV.csv");
        file.createNewFile();
        
        writer = new FileWriter(file2);
                
        while (scan.hasNext()) {
              
           String csv = scan.nextLine().replace(" ", ",");
             /**if( csv.length() < 15) {
                writer.append(" ** ");
                scan.reset();
                continue;
            }**/
            System.out.println(csv);
            System.out.println("Length: " + csv.length());
            writer.append(csv);
            writer.append("\n");
            writer.flush(); 
        }
        file.delete();
        getData(size, templatePath);
     }

public static void getData(int size, String templatePath) throws FileNotFoundException, IOException{
    System.out.println("******************************");
    String stage = null;
    String a = null;
    String Be= null;
    String Na= null;
    String Mg = null;
    String Al= null;
    String K= null;
    String Ca= null;
    String Ti= null;
    String Cr= null;
    String Mn= null;
    String Fe= null;
    String Co= null;
    String Ni= null;
    String Cu= null;
    String Ga= null;
    String Zr= null;
    String Mo= null;
    String Ru= null;
    String Cd= null;
    String In= null;
    String Sn= null;
    String Li= null;
    String Zn= null;
    String Sb= null;
    String W= null;
    String Pb= null;
    String material = null;
    String lotNum = null;
    String analyte = null;
    String lot = null;
    
    List<String> list = new ArrayList<String>();
    List<String> Ion = new ArrayList<String>();
    
    FileResource csv = new FileResource ("C:\\PDFTester\\CSV.csv");
    CSVParser parser = csv.getCSVParser(false);
    for (CSVRecord record : parser) {
        
        a = record.get(0);
        if (a.contains("Material:")){
           System.out.println(a + " " + record.get(1));
           material = record.get(1); 
           Ion.add(record.get(0));
           list.add(material);
         }
        
        if (a.contains("Lot")) {
            System.out.println(a + "" + record.get(2));
            lot = record.get(2);
            list.add(lot);
            Ion.add(record.get(0));
            lotNum = record.get(2);
        }
        if (a.contains("Stage")) {
            System.out.println(a + "" + record.get(2));
            stage = record.get(2);
            list.add(stage);
            Ion.add(record.get(0));
        }
        if (a.contains("Analyte")) {
            System.out.println(a + "" + record.get(6));
            analyte = record.get(6);
            list.add(analyte);
            Ion.add(record.get(0));
        }
        if(a.contains("Be") && a.length() <= 3){
            Be = record.get(3);
            list.add(Be);
            Ion.add(record.get(0));
        }
                
        if( a.contains("Na") && a.length() <= 3) {
            Na = record.get(3);
            list.add(Na);
            Ion.add(record.get(0));
        }
        if (a.contains("Mg") && a.length() <= 3){
            Mg = record.get(3);
            list.add(Mg);
            Ion.add(record.get(0));
        }
        if(a.contains("Al") && a.length() <= 3){
            Al = record.get(3);
            list.add(Al);
            Ion.add(record.get(0));
        }
        if(a.contains("K") && a.length() <= 3) {
            K = record.get(3);
            list.add(K);
            Ion.add(record.get(0));
        }
        if(a.contains("Ca") && a.length() <= 3){ 
            Ca = record.get(3);
            list.add(Ca);
            Ion.add(record.get(0));
        }
        if(a.contains("Ti") && a.length() <= 3) {
            Ti = record.get(3);
            list.add(Ti);
            Ion.add(record.get(0));
        }
        if(a.contains("Cr") && a.length() <= 3) {
            Cr = record.get(3);
            list.add(Cr);
            Ion.add(record.get(0));
        }
        if(a.contains("Mn") && a.length() <= 3) {
           Mn = record.get(3);
           list.add(Mn);
            Ion.add(record.get(0));
        }
        if(a.contains("Fe") && a.length() <= 3) {
            Fe = record.get(3);
            list.add(Fe);
            Ion.add(record.get(0));
        }
        if(a.contains("Co") && a.length() <= 3){
            Co = record.get(3);
            list.add(Co);
            Ion.add(record.get(0));
        }
        if(a.contains("Ni") && a.length() <= 3){
            Ni = record.get(3);
            list.add(Ni);
            Ion.add(record.get(0));
        }
        if(a.contains("Cu") && a.length() <= 3){
            Cu = record.get(3);
            list.add(Cu);
            Ion.add(record.get(0));
        }
        if(a.contains("Ga") && a.length() <= 3){
            Ga = record.get(3);
            list.add(Ga);
            Ion.add(record.get(0));
        }
        if(a.contains("Zr") && a.length() <= 3){
            Zr = record.get(3);
            list.add(Zr);
            Ion.add(record.get(0));
        }
        if(a.contains("Mo") && a.length() <= 3){
            Mo = record.get(3);
            list.add(Mo);
            Ion.add(record.get(0));
        }
        if(a.contains("Ru") && a.length() <= 3){
            Ru = record.get(3);
            list.add(Ru);
            Ion.add(record.get(0));
        }
        if(a.contains("Cd") && a.length() <= 3){
            Cd = record.get(3);
            list.add(Cd);
            Ion.add(record.get(0));
        }
        if(a.contains("In") && a.length() <= 3){
            In = record.get(3);
            list.add(In);
            Ion.add(record.get(0));
        }
        if(a.contains("Sn") && a.length() <= 3){
            Sn = record.get(3);
            list.add(Sn);
            Ion.add(record.get(0));
        }
        if(a.contains("Li") && a.length() <= 3){
            Li = record.get(3);
            list.add(Li);
            Ion.add(record.get(0));
        }
        if(a.contains("Zn") && a.length() <= 3){
            Zn = record.get(3);
            list.add(Zn);
            Ion.add(record.get(0));
        }
        if(a.contains("Sb") && a.length() <= 3){
            Sb = record.get(3);
            list.add(Sb);
            Ion.add(record.get(0));
        }
        if(a.contains("W") && a.length() <= 3){
            W = record.get(3);
            list.add(W);
            Ion.add(record.get(0));
        }
        if(a.contains("Pb") && a.length() <= 3){
            Pb = record.get(3);
            list.add(Pb);
            Ion.add(record.get(0));
        }
    }
         
       
       addToExcel(list, Ion, material, lotNum, size, templatePath);
}  
//public static void excelTemplate(List list, List Ion, String material, String lotNum, int size) throws IOException{
public static void excelTemplate(List selectedFiles, int size){   
    System.out.println("***********************");
    System.out.println("Creating Excel Template");
    System.out.println("***********************");

    DateFormat dateFormat = new SimpleDateFormat("_yyyy_MM_dd_HH:mm:ss");
    Date date = new Date();
    String templatePath = null; 
    try { 
        File template = new File("C:\\PDFTester\\_Ions.xls");
        template.createNewFile();
        try (FileOutputStream ions = new FileOutputStream(template, false)) {
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet worksheet = workbook.createSheet("Ions");

            System.out.println("Populating common fields..... ");
            
            System.out.println(selectedFileSize);
            
            HSSFRow name = worksheet.createRow((short) 0);
            HSSFCell cellA1 = name.createCell((short) 0);
            cellA1.setCellValue("Name: ");
            HSSFCellStyle cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            
            
            HSSFRow lot = worksheet.createRow((short) 1);
            HSSFCell cellA2 = lot.createCell((short) 0);
            cellA2.setCellValue("Lot #: ");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
           
            
            HSSFRow stage = worksheet.createRow((short) 2);
            HSSFCell cellA3 = stage.createCell((short) 0);
            cellA3.setCellValue("Stage: ");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            
            
            HSSFRow conc = worksheet.createRow((short) 3);
            HSSFCell cellA4 = conc.createCell((short) 0);
            cellA4.setCellValue("Analyte");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Be = worksheet.createRow((short) 4);
            HSSFCell cellA5= Be.createCell((short) 0);
            cellA5.setCellValue("Be");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Na = worksheet.createRow((short) 5);
            HSSFCell cellA6= Na.createCell((short) 0);
            cellA6.setCellValue("Na");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Mg = worksheet.createRow((short) 6);
            HSSFCell cellA7= Mg.createCell((short) 0);
            cellA7.setCellValue("Mg");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Al = worksheet.createRow((short) 7);
            HSSFCell cellA8= Al.createCell((short) 0);
            cellA8.setCellValue("Al");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow K = worksheet.createRow((short) 8);
            HSSFCell cellA9= K.createCell((short) 0);
            cellA9.setCellValue("K");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Ca = worksheet.createRow((short) 9);
            HSSFCell cellA10= Ca.createCell((short) 0);
            cellA10.setCellValue("Ca");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Ti = worksheet.createRow((short) 10);
            HSSFCell cellA11= Ti.createCell((short) 0);
            cellA11.setCellValue("Ti");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Cr = worksheet.createRow((short) 11);
            HSSFCell cellA12= Cr.createCell((short) 0);
            cellA12.setCellValue("Cr");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Mn = worksheet.createRow((short) 12);
            HSSFCell cellA13= Mn.createCell((short) 0);
            cellA13.setCellValue("Mn");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Fe = worksheet.createRow((short) 13);
            HSSFCell cellA14= Fe.createCell((short) 0);
            cellA14.setCellValue("Fe");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Ni = worksheet.createRow((short) 15);
            HSSFCell cellA16= Ni.createCell((short) 0);
            cellA16.setCellValue("Ni");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Co = worksheet.createRow((short) 14);
            HSSFCell cellA15= Co.createCell((short) 0);
            cellA15.setCellValue("Co");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Cu = worksheet.createRow((short) 16);
            HSSFCell cellA17= Cu.createCell((short) 0);
            cellA17.setCellValue("Cu");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Ga = worksheet.createRow((short) 17);
            HSSFCell cellA18= Ga.createCell((short) 0);
            cellA18.setCellValue("Ga");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Zr = worksheet.createRow((short) 18);
            HSSFCell cellA19= Zr.createCell((short) 0);
            cellA19.setCellValue("Zr");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Mo = worksheet.createRow((short) 19);
            HSSFCell cellA20= Mo.createCell((short) 0);
            cellA20.setCellValue("Mo");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Ru = worksheet.createRow((short) 20);
            HSSFCell cellA21= Ru.createCell((short) 0);
            cellA21.setCellValue("Ru");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Cd = worksheet.createRow((short) 21);
            HSSFCell cellA22= Cd.createCell((short) 0);
            cellA22.setCellValue("Cd");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow In = worksheet.createRow((short) 22);
            HSSFCell cellA23= In.createCell((short) 0);
            cellA23.setCellValue("In");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Sn = worksheet.createRow((short) 23);
            HSSFCell cellA24= Sn.createCell((short) 0);
            cellA24.setCellValue("Sn");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Li = worksheet.createRow((short) 24);
            HSSFCell cellA25= Li.createCell((short) 0);
            cellA25.setCellValue("Li");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Zn = worksheet.createRow((short) 25);
            HSSFCell cellA26= Zn.createCell((short) 0);
            cellA26.setCellValue("Zn");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Sb = worksheet.createRow((short) 26);
            HSSFCell cellA27= Sb.createCell((short) 0);
            cellA27.setCellValue("Sb");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow W = worksheet.createRow((short) 27);
            HSSFCell cellA28= W.createCell((short) 0);
            cellA28.setCellValue("W");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow Pb = worksheet.createRow((short) 28);
            HSSFCell cellA29= Pb.createCell((short) 0);
            cellA29.setCellValue("Pb");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow tot = worksheet.createRow((short) 29);
            HSSFCell cellA30= tot.createCell((short) 0);
            cellA30.setCellValue("Total: ");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow row30 = worksheet.createRow((short) 30);
            HSSFCell cellA31= tot.createCell((short) 0);
            
            HSSFRow row31 = worksheet.createRow((short) 31);
            HSSFCell cellA32= tot.createCell((short) 0);
            
            HSSFRow critHeader = worksheet.createRow((short) 32);
            HSSFCell cellA33= critHeader.createCell((short) 0);
            cellA33.setCellValue("10 Critical Ions");
            worksheet.addMergedRegion(new CellRangeAddress(32,32,0,size));
            CellUtil.setAlignment(cellA33,workbook, cellStyle.ALIGN_CENTER);
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critLot = worksheet.createRow((short) 33);
            HSSFCell cellA34= critLot.createCell((short) 0);
            cellA34.setCellValue("Lot: ");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critConc = worksheet.createRow((short) 34);
            HSSFCell cellA35= critConc.createCell((short) 0);
            cellA35.setCellValue("Analyte");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critNa = worksheet.createRow((short) 35);
            HSSFCell cellA36= critNa.createCell((short) 0);
            cellA36.setCellValue("Na");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critMg = worksheet.createRow((short) 36);
            HSSFCell cellA37= critMg.createCell((short) 0);
            cellA37.setCellValue("Mg");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critAl = worksheet.createRow((short) 37);
            HSSFCell cellA38= critAl.createCell((short) 0);
            cellA38.setCellValue("Al");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critK = worksheet.createRow((short) 38);
            HSSFCell cellA39= critK.createCell((short) 0);
            cellA39.setCellValue("K");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critCa = worksheet.createRow((short) 39);
            HSSFCell cellA40= critCa.createCell((short) 0);
            cellA40.setCellValue("Ca");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critCr = worksheet.createRow((short) 40);
            HSSFCell cellA41= critCr.createCell((short) 0);
            cellA41.setCellValue("Cr");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critMn = worksheet.createRow((short) 41);
            HSSFCell cellA42= critMn.createCell((short) 0);
            cellA42.setCellValue("Cr");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critFe = worksheet.createRow((short) 42);
            HSSFCell cellA43= critFe.createCell((short) 0);
            cellA43.setCellValue("Fe");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critNi = worksheet.createRow((short) 43);
            HSSFCell cellA44= critNi.createCell((short) 0);
            cellA44.setCellValue("Cr");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critCu = worksheet.createRow((short) 44);
            HSSFCell cellA45= critCu.createCell((short) 0);
            cellA45.setCellValue("Cu");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            
            HSSFRow critTot = worksheet.createRow((short) 45);
            HSSFCell cellA46= critTot.createCell((short) 0);
            cellA46.setCellValue("Total: ");
            cellStyle = workbook.createCellStyle();
            cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
            cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
            //****************************************************************************************************************             
            
            
            
            workbook.write(ions);
            ions.flush();
        }
            templatePath = template.getAbsolutePath();
            PDFGui.go(templatePath, selectedFiles, size);
            //addToExcel(list, Ion, material, lotNum, size, path);
           } catch (IOException ex) {
        Logger.getLogger(PDFTest.class.getName()).log(Level.SEVERE, null, ex);
    }
        
    }

public static void addToExcel(List list, List Ion, String material, String lotNum, int size, String templatePath) throws FileNotFoundException, IOException{  
    System.out.println("***************************************************");
    System.out.println("Starting AddToExcel");
    System.out.println("***************************************************");
    String[] listString = (String[]) list.toArray(new String[0]); 
    String[] ionString = (String[]) Ion.toArray(new String[0]);
    System.out.println("List: " + list);
    System.out.println("Ions: " + Ion);
    
    String nameLot = listString[1].substring(0, 7); 
    
            FileInputStream template = new FileInputStream(new File(templatePath));
            
            HSSFWorkbook workbook = new HSSFWorkbook(template);
            HSSFSheet worksheet = workbook.getSheetAt(0);
            
            System.out.println("n = " + n);
            
            System.out.println("Files to work with: ");
            System.out.print(selectedFileSize + ", on number " + n);
            HSSFCellStyle cellStyle = workbook.createCellStyle();
            
            HSSFRow name = worksheet.getRow((short) 0);
                HSSFCell cellA1 = name.getCell((short) 0);
                
               
             HSSFRow lot = worksheet.getRow((short) 1);
                HSSFCell cellA2 = name.getCell((short) 0);
                

            HSSFRow stage = worksheet.getRow((short) 2);
                HSSFCell cellA3 = name.getCell((short) 0);
                
                
            HSSFRow conc = worksheet.getRow((short) 3);    
                HSSFCell cellA4 = conc.getCell((short) 0);
                
            HSSFRow Be = worksheet.getRow((short) 4);
                HSSFCell cellA5= Be.getCell((short) 0);
                
            HSSFRow Na = worksheet.getRow((short) 5);
                HSSFCell cellA6= Na.getCell((short) 0);
                
            HSSFRow Mg = worksheet.getRow((short) 6);
                HSSFCell cellA7= Mg.getCell((short) 0);
                
            HSSFRow Al = worksheet.getRow((short) 7);
                HSSFCell cellA8= Al.getCell((short) 0);
               
            HSSFRow K = worksheet.getRow((short) 8);
                HSSFCell cellA9= K.getCell((short) 0);
                
            HSSFRow Ca = worksheet.getRow((short) 9);
                HSSFCell cellA10= Ca.getCell((short) 0);
               
            HSSFRow Ti = worksheet.getRow((short) 10);
                HSSFCell cellA11= Ti.getCell((short) 0);
               
            HSSFRow Cr = worksheet.getRow((short) 11);
                HSSFCell cellA12= Cr.getCell((short) 0);
                
            HSSFRow Mn = worksheet.getRow((short) 12);
                HSSFCell cellA13= Mn.getCell((short) 0);
               
            HSSFRow Fe = worksheet.getRow((short) 13);
                HSSFCell cellA14= Fe.getCell((short) 0);
             
            HSSFRow Co = worksheet.getRow((short) 14);
                HSSFCell cellA15= Co.getCell((short) 0);   
                
            HSSFRow Ni = worksheet.getRow((short) 15);
                HSSFCell cellA16= Ni.getCell((short) 0);
            
            HSSFRow Cu = worksheet.getRow((short) 16);
                HSSFCell cellA17= Cu.getCell((short) 0);
                
            HSSFRow Ga = worksheet.getRow((short) 17);
                HSSFCell cellA18= Ga.getCell((short) 0);
                
            HSSFRow Zr = worksheet.getRow((short) 18);
                HSSFCell cellA19= Zr.getCell((short) 0);
                
            HSSFRow Mo = worksheet.getRow((short) 19);
                HSSFCell cellA20= Mo.getCell((short) 0);
               
            HSSFRow Ru = worksheet.getRow((short) 20);
                HSSFCell cellA21= Ru.getCell((short) 0);
                
            HSSFRow Cd = worksheet.getRow((short) 21);
                HSSFCell cellA22= Cd.getCell((short) 0);
              
            HSSFRow In = worksheet.getRow((short) 22);
                HSSFCell cellA23= In.getCell((short) 0);
               
            HSSFRow Sn = worksheet.getRow((short) 23);
                HSSFCell cellA24= Sn.getCell((short) 0);
               
            HSSFRow Li = worksheet.getRow((short) 24);
                HSSFCell cellA25= Li.getCell((short) 0);
              
            HSSFRow Zn = worksheet.getRow((short) 25);
                HSSFCell cellA26= Zn.getCell((short) 0);
               
            HSSFRow Sb = worksheet.getRow((short) 26);
                HSSFCell cellA27= Sb.getCell((short) 0);
              
            HSSFRow W = worksheet.getRow((short) 27);
                HSSFCell cellA28= W.getCell((short) 0);
                
             HSSFRow Pb = worksheet.getRow((short) 28);
                HSSFCell cellA29= Pb.getCell((short) 0);
                
            HSSFRow tot = worksheet.getRow((short) 29);
                HSSFCell cellA30= tot.getCell((short) 0);
            
            HSSFRow row31 = worksheet.getRow((short) 30);
                HSSFCell cellA31= row31.getCell((short) 0); 
                
            HSSFRow row32 = worksheet.getRow((short) 31);
                HSSFCell cellA32= row31.getCell((short) 0); 
                
                System.out.println("Now starting Crit Ions");
                
            HSSFRow critHeader = worksheet.getRow((short) 32);
                HSSFCell cellA33= critHeader.getCell((short) 0);
                
            HSSFRow critLot = worksheet.getRow((short) 33);
                HSSFCell cellA34= critLot.getCell((short) 0);
            
            HSSFRow critConc = worksheet.getRow((short) 34);
                HSSFCell cellA35= critConc.getCell((short) 0);
                
            HSSFRow critNa = worksheet.getRow((short) 35);
                HSSFCell cellA36= critNa.getCell((short) 0);
                
            HSSFRow critMg = worksheet.getRow((short) 36);
                HSSFCell cellA37= critMg.getCell((short) 0);
                
            HSSFRow critAl = worksheet.getRow((short) 37);
                HSSFCell cellA38= critAl.getCell((short) 0);
            
            HSSFRow critK = worksheet.getRow((short) 38);
                HSSFCell cellA39= critK.getCell((short) 0);
                
            HSSFRow critCa = worksheet.getRow((short) 39);
                HSSFCell cellA40= critCa.getCell((short) 0);
                
            HSSFRow critCr = worksheet.getRow((short) 40);
                HSSFCell cellA41= critCr.getCell((short) 0);
            
            HSSFRow critMn = worksheet.getRow((short) 41);
                HSSFCell cellA42= critMn.getCell((short) 0);
                
            HSSFRow critFe = worksheet.getRow((short) 42);
                HSSFCell cellA43= critFe.getCell((short) 0);
                
            HSSFRow critNi = worksheet.getRow((short) 43);
                HSSFCell cellA44= critNi.getCell((short) 0);
                
            HSSFRow critCu = worksheet.getRow((short) 44);
                HSSFCell cellA45= critCu.getCell((short) 0);
                
            HSSFRow critTot = worksheet.getRow((short) 45);
                HSSFCell cellA46= critTot.getCell((short) 0);
                
            System.out.println("Now filling data from the pdf");
            System.out.println("Size: " + size);
             
          int totalRowNum = worksheet.getPhysicalNumberOfRows() + 3;  
          System.out.println("Last row at: " + totalRowNum);
            

            for (int r = 0; r < size; r++){ // check each row
               if(n == size){
                   break;
               }
                   Row rw = worksheet.getRow(r);
               
               System.out.println("Row Number: " + (r + 1));
               if(rw == null) {
                   System.out.println("Row ERROR");
                   continue;
                    }
               System.out.println("No Row errors: ");
                for (int x = 0; x < totalRowNum; x++){ //check each cell
                    
                     Cell c = rw.getCell(x);
                   
                    if(c == null){    //if cell is null, make it Blank
                        c = rw.getCell(x, Row.CREATE_NULL_AS_BLANK);
                System.out.println("Converting Null cells to Blank");
                    }
                    
                     System.out.println("No Cell Errors");
                     
                     if (c.getCellType() == Cell.CELL_TYPE_BLANK){ //if cell is blank
                         System.out.println("No more Null Cells");
                              //fill blank cell
                                int i = 0;
                                System.out.println("Populating spreadsheet... " + n);
            
            // index from 0,0... cell A1 is cell(0,0)
                HSSFCell cellB1 = name.createCell((short) n);
                cellB1.setCellValue(listString[i]);
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.GOLD.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
                HSSFCell cellB2 = lot.createCell((short) n);
                cellB2.setCellValue(listString[i]);
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                 i++;
                 
                HSSFCell cellB3 = stage.createCell((short) n);
                cellB3.setCellValue(listString[i]);
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                 
                HSSFCell cellB4 = conc.createCell((short) n);
                cellB4.setCellValue("Conc. (ppb)");
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
            
                
                
                HSSFCell cellB5 = Be.createCell((short) n);
                cellB5.setCellValue(Double.parseDouble(listString[i]));
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.GREEN.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
    
                HSSFCell cellB6 = Na.createCell((short) n);
                cellB6.setCellValue(Double.parseDouble(listString[i]));
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB7 = Mg.createCell((short) n);
                cellB7.setCellValue(Double.parseDouble(listString[i]));
                cellStyle = workbook.createCellStyle();
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
               
                HSSFCell cellB8 = Al.createCell((short) n);
                cellB8.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB9 = K.createCell((short) n);
                cellB9.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
            
                HSSFCell cellB10 = Ca.createCell((short) n);
                cellB10.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
                
                
                HSSFCell cellB11 = Ti.createCell((short) n);
                cellB11.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                HSSFCell cellB12 = Cr.createCell((short) n);
                cellB12.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB13 = Mn.createCell((short) n);
                cellB13.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB14 = Fe.createCell((short) n);
                cellB14.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB15 = Co.createCell((short) n);
                cellB15.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB16 = Ni.createCell((short) n);
                cellB16.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                HSSFCell cellB17 = Cu.createCell((short) n);
                cellB17.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                i++;
                
            
                
                
                HSSFCell cellB18 = Ga.createCell((short) n);
                cellB18.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
               
                
                i++;
                
                HSSFCell cellB19 = Zr.createCell((short) n);
                cellB19.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
             
                
                i++;
                
                HSSFCell cellB20 = Mo.createCell((short) n);
                cellB20.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
            
                
                i++;
                
                HSSFCell cellB21 = Ru.createCell((short) n);
                cellB21.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
            
                i++;
                
                HSSFCell cellB22 = Cd.createCell((short) n);
                cellB22.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                
            
                i++;
                
                HSSFCell cellB23 = In.createCell((short) n);
                cellB23.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                
            
                i++;
                
                HSSFCell cellB24 = Sn.createCell((short) n);
                cellB24.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                
            
                i++;
                
                HSSFCell cellB25 = Li.createCell((short) n);
                cellB25.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                
            
                i++;
                
                HSSFCell cellB26 = Zn.createCell((short) n);
                cellB26.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                
            
                i++;
                
                HSSFCell cellB27 = Sb.createCell((short) n);
                cellB27.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
              
            
                i++;
                
                HSSFCell cellB28 = W.createCell((short) n);
                cellB28.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);

                i++;
                
                HSSFCell cellB29= Pb.createCell((short) n);
                cellB29.setCellValue(Double.parseDouble(listString[i]));
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB30 = tot.createCell((short) n);
                cellB30.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String sum = "SUM(" + alphabet[n]+"5:"+ alphabet[n]+"29)"; 
                cellB30.setCellFormula(sum);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB34 = critLot.createCell((short) n);
                cellB34.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critLotVal =  (alphabet[n]+ "3"); 
                cellB34.setCellFormula(critLotVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB35 = critConc.createCell((short) n);
                cellB35.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critConcVal =  (alphabet[n]+ "4"); 
                cellB35.setCellFormula(critConcVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB36 = critNa.createCell((short) n);
                cellB36.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critNaVal =  (alphabet[n]+ "6"); 
                cellB36.setCellFormula(critNaVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB37 = critMg.createCell((short) n);
                cellB37.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critMgVal =  (alphabet[n]+ "7"); 
                cellB37.setCellFormula(critMgVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB38 = critAl.createCell((short) n);
                cellB38.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critAlVal =  (alphabet[n]+ "8"); 
                cellB38.setCellFormula(critAlVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB39 = critK.createCell((short) n);
                cellB39.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critKVal = ( alphabet[n]+ "9"); 
                cellB39.setCellFormula(critKVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB40 = critCa.createCell((short) n);
                cellB40.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critCaVal =  (alphabet[n]+ "10"); 
                cellB40.setCellFormula(critCaVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB41 = critCr.createCell((short) n);
                cellB41.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critCrVal =  (alphabet[n]+ "12"); 
                cellB41.setCellFormula(critCrVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB42 = critMn.createCell((short) n);
                cellB42.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critMnVal =  (alphabet[n]+ "13"); 
                cellB42.setCellFormula(critMnVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB43 = critFe.createCell((short) n);
                cellB43.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critFeVal =  (alphabet[n]+ "14"); 
                cellB43.setCellFormula(critFeVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB44 = critNi.createCell((short) n);
                cellB44.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critNiVal = ( alphabet[n] + "16"); 
                cellB44.setCellFormula(critNiVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB45 = critCu.createCell((short) n);
                cellB45.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critCuVal =  (alphabet[n]+ "17"); 
                cellB45.setCellFormula(critCuVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                
                HSSFCell cellB46 = critTot.createCell((short) n);
                cellB46.setCellType(HSSFCell.CELL_TYPE_FORMULA);
                String critTotVal =  ("SUM(" + alphabet[n]+"36:"+ alphabet[n]+"45)"); 
                cellB46.setCellFormula(critTotVal);
                cellStyle.setFillForegroundColor(HSSFColor.LIGHT_CORNFLOWER_BLUE.index);
                cellStyle.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                                
                           }

                      }
             
                }
     FileOutputStream fileOut = new FileOutputStream(templatePath);
        workbook.write(fileOut);
        fileOut.flush();
        
            n++;
        String destinationPath = material + "_" + nameLot + "_Ions.xls";
          
          
            System.out.println("*************Template filled*************");
            System.out.println("Now renaming file for you");
            
           
                copyAndClean(templatePath, destinationPath);
            
            
    }
    public static void copyAndClean(String sourcePath, String destinationPath) throws IOException {
        
        Files.copy(Paths.get(sourcePath), new FileOutputStream(destinationPath)); //saves to unique output file
        System.out.println("Your spreadsheet is located at: " + destinationPath);
         System.out.println("****************COMPLETE****************");
    } 
}   
            

    




