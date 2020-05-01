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
    public static void main(String[] args) {
        ReadWordFile wordReader = new ReadWordFile("D:\\TDT\\HK6_2020\\Software Testing\\MidTerm Project\\TieuLuan_Nhom3_UnitTest.docx");
        // wordReader.write();
        wordReader.writePictures();
        
        //ReadExcelFile excelReader = new ReadExcelFile("C:\\Users\\doduy\\Downloads\\Customer.xlsx");
        //excelReader.print();
        //excelReader.writePictures();
        //excelReader.write();

        // ReadPdfFile pdfReader = new ReadPdfFile("C:\\Users\\doduy\\Downloads\\Documents\\Seminar_08.pdf"); // ĐƯờng dẫn được copy từ windows explorer
        // pdfReader.writePictures(); // Lưu ảnh từ pdf
        // pdfReader.write();
    }
}
