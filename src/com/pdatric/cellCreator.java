/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.pdatric;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

/**
 *
 * @author pluebbert
 */
public class cellCreator {
    private HSSFRow row;
    private HSSFCell cell; 
    private String[] nameString = {"name", "lot", "stage", "conc", "Be", "Na", "Mg", 
                            "Al", "K", "Ca", "Ti", "Cr", "Mn", "Fe", "Co", "Ni",
                            "Cu", "Ga", "Zr", "Mo", "Ru", "Cd", "In", "Sn", "Li",
                            "Zn", "Sb", "W", "Pb", "row30", "row31", "critHeader",
                            "critLot", "critConc", "critNa", "critMg", "critAl", 
                            "critK", "critCa", "critCr", "critMn", "critFe",
                            "critNi", "critCu", "critTot"};
    private HSSFWorkbook workbook;
    private HSSFSheet sheet;

    public cellCreator(HSSFSheet worksheet) {
        this.sheet = worksheet;
        
    }
    
    public String[] getNameString(){
        return nameString;
    }


    public HSSFRow getRow() {
        return row;
    }

    public void setRow(HSSFRow row) {
        this.row = row;
    }

    public HSSFCell getCell() {
        return cell;
    }

    public void setCell(HSSFCell cell) {
        this.cell = cell;
    }

    public HSSFWorkbook getWorkbook() {
        return workbook;
    }

    public HSSFSheet getWorksheet() {
        return sheet;
    }
    
    
    public void createCells(String[] text, HSSFCellStyle style){
        System.out.println("Using CEllCREATOR class");
        System.out.println("text.lengths: " + text.length);
        for(int i = 0; i < text.length; i++){
            if(i == 29 || i == 30){
                row = sheet.createRow(i);
                cell = row.createCell(0);
            }
            else{
                row = sheet.createRow(i); 
                cell = row.createCell(0);
                cell.setCellValue(text[i]);
                cell.setCellStyle(style);
            }
        }
    }
    
    public void getRow(HSSFSheet sheet, int i){
        
    }
    
}
