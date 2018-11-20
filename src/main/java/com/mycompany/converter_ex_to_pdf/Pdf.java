
package com.mycompany.converter_ex_to_pdf;

import com.itextpdf.text.BaseColor;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Font;
import com.itextpdf.text.PageSize;
import com.itextpdf.text.Phrase;
import com.itextpdf.text.pdf.PdfPCell;
import com.itextpdf.text.pdf.PdfPTable;
import com.itextpdf.text.pdf.PdfWriter;
import java.awt.Color;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.util.Iterator;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;


public class Pdf {
    
    /**
     * Create a new PDF file
     * 
     * @param path Save PDF file path
     * @return 
     */
    public Document createNewPDF(String path){
        Document pdfDoc = new Document(PageSize.A4, 10f, 10f, 10f, 10f);
        try {
            PdfWriter.getInstance(pdfDoc, new FileOutputStream(new File(path)));
            pdfDoc.open();
        } catch (FileNotFoundException | DocumentException ex) {
            System.out.println("Error creating PDF file");
        }
        return pdfDoc;
    }
    
    /**
     * Create a new table to add to the PDF file
     * 
     * @param widthColumns
     * @return PdfTable The object of the table
     */
    public PdfPTable createNewTable(float[] widthColumns){
        return new PdfPTable(widthColumns);
    }
    
    /**
     * Add content to the PDF table
     * 
     * @param table
     * @param rowIterator
     * @param numberColumns 
     */
    public void setContentTable(PdfPTable table, Iterator<Row> rowIterator, int numberColumns){
        PdfPCell table_cell;
        
        while(rowIterator.hasNext()) {
            int cpt = 0;
            Row row = rowIterator.next(); 
            
            Iterator<Cell> cellIterator = row.cellIterator();
            while(cellIterator.hasNext()) {
                cpt++;
                HSSFCell cell = (HSSFCell) cellIterator.next();
                cell.setCellType(CellType.STRING);
                
                int fontSize = cell.getCellStyle().getFont(cell.getSheet().getWorkbook()).getFontHeightInPoints();
                int fontColor = cell.getCellStyle().getFont(cell.getSheet().getWorkbook()).getColor();
                
                Font font = new Font(Font.getFamily("Arial"), fontSize, Font.NORMAL, new BaseColor(Color.getColor("color", fontColor)));
                table_cell = new PdfPCell(new Phrase(cell.getStringCellValue(), font));
                
                if(!cellIterator.hasNext() && cpt < numberColumns){
                    table_cell.setColspan(numberColumns);
                    cpt = numberColumns;
                    table_cell.setBorder(0);
                }
                table.addCell(table_cell);
            }
            if(cpt < numberColumns){
                
                while(cpt < numberColumns){
                    cpt++;
                    table_cell = new PdfPCell(new Phrase(""));
                    
                    table_cell.setBorder(0);
                    table.addCell(table_cell);
                    
                }
            }
        }
    }
    
    /**
     * Add a table file to a PDF file
     * 
     * @param pdfDoc PDF file
     * @param table The table to be added to the PDF file
     * @return 
     */
    public boolean addTable(Document pdfDoc, PdfPTable table){
        try {
            pdfDoc.add(table);
            return true;
        } catch (DocumentException ex) {
            System.out.println("Error : Failed adding table to PDF file");
            return false;
        }
    }
    
}
