
package com.mycompany.converter_ex_to_pdf;

import java.io.FileInputStream;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.ss.usermodel.*;
import java.util.Iterator;
import java.io.File;
import java.io.IOException;
import java.net.URL;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;


public class Excel {
    
    //Table rows of the Excel file
    private HSSFSheet worksheet;
    //The largest line number in the table
    private int lln;
    
    /**
     * Retrieve the Excel file to be converted to PDF
     * 
     * @param urlString Excel file url.
     * @param path 
     * @return FileInputStream Enroll the Excel file
     */
    public FileInputStream getExcelFile(String urlString, String path){
        
        FileInputStream excelDoc = null;
        try {
            URL url = new URL(urlString);
            Path targetPath = new File(path+"\\excel.xls").toPath();
            Files.copy(url.openStream(), targetPath, StandardCopyOption.REPLACE_EXISTING);
            excelDoc = new FileInputStream(new File(path+"\\excel.xls"));
        } catch (IOException ex) {
            System.out.println("Error : Error fetching file");
        }
        return excelDoc;
    }
    
    /**
     * Extract the table rows from the Excel file
     * 
     * @param excelDoc Enroll the Excel file
     * @return Iterator Table rows of the Excel file
     */
    public Iterator<Row> getRows(FileInputStream excelDoc){
        Iterator<Row> rowIterator = null;
        try {
            if(worksheet == null){
                HSSFWorkbook xls_workbook = new HSSFWorkbook(excelDoc);
                worksheet = xls_workbook.getSheetAt(0);
            }
            rowIterator = worksheet.iterator();
            
        } catch (IOException ex) {
            System.out.println(ex.getMessage());
        }
        
        return rowIterator;
    }
    
    /**
     * Extract the number of columns from the table
     * 
     * @param rowIterator Table rows of the Excel file
     * @return Integer Number. number of columns from the table
     */
    public int getNumberColumns(Iterator<Row> rowIterator){
        int rowNumber = 0;
        int numberColumns = 0;
        int cpt;
        lln = 0;
        while(rowIterator.hasNext()){
            rowNumber++;
            Row row = rowIterator.next();
            Iterator<Cell> cellIterator = row.cellIterator();
            cpt = 0;
            while(cellIterator.hasNext()) {
                cellIterator.next();
                cpt++;
            }
            if(cpt > numberColumns){ 
                numberColumns = cpt;
                lln = rowNumber;
            }
        }
        return numberColumns;
    }
    
     /**
     * Extract the width of each cell from the table
     * 
     * @param rowIterator Table rows of the Excel file
     * @param numberColumns Number of column in Excel file
     * @return A float number matrix represents the width of each cell
     */
    public float[] getWidthsColumns(Iterator<Row> rowIterator, int numberColumns){
        
        float[] widths = new float[numberColumns];
        if(rowIterator.hasNext()){
            Row row = rowIterator.next();
            for(int i =0; i< lln-1; i++){
                row = rowIterator.next();
            }
            
            Iterator<Cell> cellIterator = row.cellIterator();
            for(int  i =0; i< numberColumns; i++){
                Cell cell = cellIterator.next();
                widths[i] = cell.getSheet().getColumnWidth(i);
            }
        }
        
        return widths;
    }
    
}
