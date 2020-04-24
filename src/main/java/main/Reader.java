/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import reader.ReadWordFile;
import reader.ReadPdfFile;
import reader.ReadExcelFile;

/**
 *
 * @author doduy
 */
public class Reader {
    ReadWordFile wordReader = new ReadWordFile("C:\\Users\\doduy\\Downloads\\Seminar_08.docx");
    

    //ReadExcelFile excelReader = new ReadExcelFile("C:\\Users\\doduy\\Downloads\\Customer.xlsx");
    //excelReader.print();
    //excelReader.writePictures();
    //excelReader.write();

    // ReadPdfFile pdfReader = new ReadPdfFile("C:\\Users\\doduy\\Downloads\\Lab01.pdf"); // ĐƯờng dẫn được copy từ windows explorer
    // pdfReader.writePictures(); // Lưu ảnh từ pdf
    // pdfReader.write(); 
}
