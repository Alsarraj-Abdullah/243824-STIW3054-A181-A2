
package com.mycompany.converter_ex_to_pdf;

import com.itextpdf.text.Document;
import com.itextpdf.text.pdf.PdfPTable;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Date;
import java.util.Iterator;
import java.util.Scanner;
import javax.swing.filechooser.FileSystemView;
import org.apache.poi.ss.usermodel.Row;


public class RunApp {
    
    Scanner scanner = new Scanner(System.in);
    
    Excel excel = null;
    Pdf pdf = null;
    
    /**
     * 
     * @param pdf
     * @param excel 
     */
    public void setObject(Pdf pdf, Excel excel){
        this.pdf = pdf;
        this.excel = excel;
    }
    
    public void start() throws IOException{
        
        String DesktopPath = System.getProperty("user.home") + "\\Desktop\\";
        String DocumentsPath =FileSystemView.getFileSystemView().getDefaultDirectory().getPath() + "\\";
        String ApplicationPath = System.getProperty("user.dir") + "\\";
        Date date= new Date();
        long timestamp = date.getTime();
        
        System.out.println("Where do you want to save " + timestamp +"-NewPDF.pdf");
                System.out.println("1.Desktop");
                System.out.println("2.My Documents");
                System.out.println("3.Application Folder");
                System.out.print("Please Enter your option : ");
        String path = getPath(scanner.nextInt());
        
        System.out.print("Enter the link of the Excel file to convert: ");
        String excelLink = scanner.next();
        
        System.out.println("download the Excel file ...");
        try (FileInputStream excelFile = excel.getExcelFile(excelLink, path)) {
            System.out.println("Extract rows from the Excel file ..");
            Iterator<Row> rows = excel.getRows(excelFile);
            
            System.out.println("Extract the number of columns from the file..");
            int numberColumns = excel.getNumberColumns(rows);
            
            rows = excel.getRows(excelFile);
            float[] widthColumns = excel.getWidthsColumns(rows, numberColumns);
            
            System.out.println("Create a new PDF file..");
            Document pdfDoc = pdf.createNewPDF(path+"\\"+ timestamp +"-NewPDF.pdf");
            
            PdfPTable table = pdf.createNewTable(widthColumns);
            
            rows = excel.getRows(excelFile);
            
            pdf.setContentTable(table, rows, numberColumns);
            
            System.out.println("Add content to the PDF file..");
            if(pdf.addTable(pdfDoc, table)){
                System.out.println("Success : Conversion completed successfully!");
                System.out.println("Saved in path: '"+path+"\\"+ timestamp +"-NewPDF.pdf"+"'\n-----------------------------------");
            }
            
            //Close the Excel file and the new PDF file
            pdfDoc.close();
        }
    }

    private String getPath(int nextInt) {
        
        String DesktopPath = System.getProperty("user.home") + "\\Desktop\\";
        String DocumentsPath =FileSystemView.getFileSystemView().getDefaultDirectory().getPath() + "\\";
        String ApplicationPath = System.getProperty("user.dir") + "\\";
        
        switch(nextInt){
            case 1:
                return DesktopPath;
            case 2:
                return DocumentsPath;
            case 3:
                return ApplicationPath;
            default: 
                return DesktopPath;
        }
    }
    
}
