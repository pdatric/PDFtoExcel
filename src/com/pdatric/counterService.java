package com.pdatric;

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */

/**
 *
 * @author pluebbert
 */

import javafx.concurrent.Task;

public class counterService extends Task {
    private int n;
    private int size;
    
    public counterService(int n, int size){
        this.n = n;
        this.size = size;
    }
    
    @Override
    public Void call() {
        final int max = size;
        updateProgress(n, size);
        
        return null;
    }
}
