package com.petsquare;

import java.io.IOException;

import static java.lang.System.out;

public class Main {

    public static void main(String[] args) {

        String path = "C:\\Users\\Asus\\Documents\\PetSquare\\DOG_WALKERS.xlsx";

        FileProcessor fp = new FileProcessor(path);

        try {
            fp.saveExcelFile(fp.validateExcelFile(fp.openExcelFile()));
        }catch (Exception e){
            out.println(e.getMessage());
            out.println(e.getStackTrace());
        }
    }
}
