/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.pdatric;

/**
 *
 * @author pluebbert
 */
public class zeroCorrection {
    private String ion;

    public zeroCorrection(String ion) {
        this.ion = ion;
    }

    public String getIon() {
        return ion;
    }

    public void setIon(String ion) {
        this.ion = ion;
    }
    
    public static String zeroCorrection(String ion){
        ion = ion.toLowerCase().replace("o", "0"); 
        if(ion.contains("—")){
            ion = ion.replace("—", "-");
            
        }
        return ion;
    }
    
    public static String minusCorrection(String ion){
        if(ion.contains("—")){
            System.out.println(ion);
            ion.replace("—", "-");
            
        }
        System.out.println("Returning : " + ion);
        return ion;
    }
    
}
