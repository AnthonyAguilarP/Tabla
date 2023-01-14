/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Class.java to edit this template
 */
package com.mycompany.tablamaven;

import java.io.FileOutputStream;
import java.util.ArrayList;
import javax.swing.JOptionPane;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author HP
 */
public class Almacen {
    public static void Excel(ArrayList<Persona> p){
        Workbook libro = new XSSFWorkbook();
        final String nombreArchivo = "Registro.xlsx";
        Sheet hoja = libro.createSheet("Hoja 1");

        String[] encabezados = {"Nombre", "Apellido","Genero"};
        int indiceFila = 0;

        Row fila = hoja.createRow(indiceFila);
        for (int i = 0; i < encabezados.length; i++) {
            String encabezado = encabezados[i];
            Cell celda = fila.createCell(i);
            celda.setCellValue(encabezado);
        }
        indiceFila++;
        for (int i = 0; i < p.size(); i++) {
            fila = hoja.createRow(indiceFila);
            fila.createCell(0).setCellValue(p.get(i).getNombre());
            fila.createCell(1).setCellValue(p.get(i).getApellido());
            fila.createCell(2).setCellValue(p.get(i).getGenero());
        indiceFila++;
        }
        FileOutputStream outputStream;
        try {
            outputStream = new FileOutputStream(nombreArchivo);
            libro.write(outputStream);
            libro.close();
            JOptionPane.showMessageDialog( null, "Exito al Guardar" );
        } catch (Exception ex) {
            JOptionPane.showMessageDialog( null, ex.getMessage() );
        }
    }
}
