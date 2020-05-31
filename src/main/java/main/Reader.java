/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package main;

import java.awt.Desktop;
import java.io.File;
import java.io.FilenameFilter;
import java.io.IOException;
import java.util.Scanner;
import reader.ReadWordFile;
import reader.ReadPdfFile;
import reader.ReadExcelFile;

/**
 *
 * @author doduy
 */
public class Reader {
    
    
    public static void main(String[] args) {
        
        Scanner sc = new Scanner(System.in);
        ReadWordFile word = null;
        ReadPdfFile pdf = null;
        ReadExcelFile excel = null;
        
        System.out.println("DEMO READ FILE MS & PDF WITH JAVA");
        System.out.print("INPUT PATH CONTAINS FILE (WORD, EXCEL, PDF): ");
        String folder = sc.nextLine();
        
        System.out.println("Root path : " + folder);
        
        // Đọc tất cả các file trong thư mục
        File folderFiles = new File(folder);
        if(folderFiles.exists()) {
            File[] files = folderFiles.listFiles(new FilenameFilter() {
                @Override
                public boolean accept(File dir, String name) {
                    if(name.startsWith("~$") || (new File(dir.getPath() + "\\" + name)).isDirectory()) {
                        return false;
                    } else {
                        return true;
                    }
                }
            });
            
            // Đọc tất cả các file trong thư mục
            for(File file : files) {
                String filename = file.getName();
                System.out.println("FILENAME: " + filename);
                
                String extension = filename.substring(filename.lastIndexOf("."), filename.length());
                
                if(extension.equals(".doc") || extension.equals(".docx")) {
                    System.out.println("Reading Word files: " + file.getPath());
                    word = new ReadWordFile(file.getPath());
                    word.write();
                    word.writePictures();
                    
                } else if (extension.equals(".xls") || extension.equals(".xlsx")) {
                    System.out.println("Reading Excel files: " + file.getPath());
                    excel = new ReadExcelFile(file.getPath());
                    excel.write();
                    excel.writePictures();
                    
                } else if (extension.equals(".pdf")) {
                    System.out.println("Reading Pdf files: " + file.getPath());
                    pdf = new ReadPdfFile(file.getPath());
                    pdf.write();
                    pdf.writePictures();
                    
                } else {
                    System.out.println("File: " + filename + " is not MS Word, Excel or PDF type");
                }
            }
            
        } else {
            System.out.println("Folder khong ton tai");
        }
        
        // Mở thư mục chứa kết quả
        open();
    }
    
    public static void open() {
        
        File output = new File(System.getProperty("user.dir") + "\\output\\");
        Desktop desktop = Desktop.getDesktop();
        try {
            
            desktop.open(output);
        } catch (IOException ex) {
            
            System.out.println("KHÔNG MỞ ĐƯỢC THƯ MỤC");
        }
    }
}
