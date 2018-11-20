
package com.mycompany.converter_ex_to_pdf;

import java.io.IOException;
import java.util.Scanner;


public class App {
    
    public static int req;
    
    public static void main(String[] args){
        
        
        Scanner scanner = new Scanner(System.in);
        
        //Create basic objects
        Pdf pdf = new Pdf();
        Excel excel = new Excel();
        
        //The program was launched
        RunApp run = new RunApp();
        run.setObject(pdf, excel);
        
        System.out.print("Press 1 to continue or 0 to exit .. ");
        req = scanner.nextInt();
        while(req == 1){
            try{
                run.start();
            }catch(IOException e){
                System.out.println("Error : Unknown Error.\n-----------------------------------");
            }
            
            System.out.print("Press 1 to restart or 0 to exit .. ");
            req = scanner.nextInt();
        }
        
    }
    
}
