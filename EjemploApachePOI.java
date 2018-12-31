/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package ejemploapachepoi;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Iterator;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author xcheko51x
 */
public class EjemploApachePOI {

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        
        //EscribirEXCEL();
        
        LeerEXCEL();
        
    }
    
    private static void LeerEXCEL() {
        
        String nombreArchivo = "ListaUsuarios.xlsx";
        String hoja = "Usuarios";
        
        try(FileInputStream file = new FileInputStream(new File(nombreArchivo))){
            //Leer archivo de Excel
            XSSFWorkbook libro = new XSSFWorkbook(file);
            // Obtener la hoja que se va a leer
            XSSFSheet sheet = libro.getSheetAt(0);
            // Obtener todas las filas de la hoja de Excel
            Iterator<Row> rowIterator = sheet.iterator();
            
            Row row;
            // Se recorre cada fila hasta el final
            while(rowIterator.hasNext()) {
                row = rowIterator.next();
                // Se obtienen las celdas por fila
                Iterator<Cell> cellIterator = row.cellIterator();
                Cell cell;
                // Se recorre cada celda
                while(cellIterator.hasNext()) {
                    // Se obtiene la celda en especifico y se imprime
                    cell = cellIterator.next();
                    System.out.print(cell.getStringCellValue()+ " - ");
                }
                System.out.println("");
            }
            
        } catch(Exception e) {
            e.getMessage();
        }
    }
    
    private static void EscribirEXCEL() {
        String nombreArchivo = "ListaUsuarios.xlsx";
        
        String hoja = "Usuarios";
        
        XSSFWorkbook libro = new XSSFWorkbook();
        XSSFSheet hoja1 = libro.createSheet(hoja);
        
        // Cabecera de la hoja de excel
        String[] header = new String[] {"NOMBRE", "TELEFONO", "EMAIL"};
        
        // Contenido de la hoja de excel
        String[][] document = new String[][] {
            {"Sergio P", "1234567", "sergiop@prueba.es"},
            {"Laura L", "4324251", "laural@prueba.es"},
            {"Juan H", "7363153", "juanh@prueba.es"}
        };
        
        // Poner en negrita la cabecera
        CellStyle style = libro.createCellStyle();
        Font font = libro.createFont();
        font.setBold(true);
        style.setFont(font);
        
        // Generar los datos para el documento
        for(int i = 0 ; i <= document.length ; i++) {
            XSSFRow row = hoja1.createRow(i); // Se crea la fila
            for(int j = 0 ; j < header.length ; j++) {
                if(i == 0) { // Para la cabecera
                    XSSFCell cell = row.createCell(j); // Se crean las celdas pra la cabecera
                    cell.setCellValue(header[j]); // Se añade el contenido
                } else {
                    XSSFCell cell = row.createCell(j); // Se crean las celdas para el contenido
                    cell.setCellValue(document[i - 1][j]); // Se añade el contenido
                }
            }
        }
        
        // Crear el archivo
        try (OutputStream fileOut = new FileOutputStream(nombreArchivo)){
            System.out.println("SE CREO EL EXCEL");
            libro.write(fileOut);
        } catch(IOException e) {
            e.printStackTrace();
        }
    }
}
