/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package reader;

import java.awt.Desktop;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStreamWriter;
import javax.imageio.ImageIO;
import org.apache.log4j.BasicConfigurator;
import org.apache.pdfbox.cos.COSName;
import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.PDPage;
import org.apache.pdfbox.pdmodel.PDPageTree;
import org.apache.pdfbox.pdmodel.PDResources;
import org.apache.pdfbox.pdmodel.graphics.PDXObject;
import org.apache.pdfbox.pdmodel.graphics.image.PDImageXObject;
import org.apache.pdfbox.text.PDFTextStripper;
/**
 *
 * @author doduy
 */
public class ReadPdfFile {
    private static final String PATH = System.getProperty("user.dir") + "\\output\\";
    private static PDDocument pdf;
    private static File rfile;
    private static File wfile;
    private static File pfile;
    private static String filename;
    private static String foldername;
    private static PDFTextStripper stripper;
    private static String folder;
    
    /**
     * Truyền vào đường dẫn + tên file được copy từ đường link trên Windows Explorer
     * @param path 
     */
    public ReadPdfFile(String path) {
        BasicConfigurator.configure();
        
        filename = path.substring(path.lastIndexOf('\\') + 1, path.length());
        foldername = filename.substring(0, filename.lastIndexOf('.'));
        
        // Tạo đối tượng ghi/đọc file nội dung Workbook
        rfile = new File(path);
        wfile = new File(PATH + foldername + "\\documents\\"); 
        pfile = new File(PATH + foldername + "\\images\\");
        
        // Tạo thư mục lưu các file ảnh
        if(!pfile.exists()) {
            pfile.mkdirs();
        }
        // Tạo thư mục lưu các sheet của workbook
        if(!wfile.exists()) {
            wfile.mkdirs();
        }
        
        try {
            pdf = PDDocument.load(rfile);
        } catch (IOException ex) {
            System.out.println("ĐỌC FILE PDF THẤT BẠI");
        }
        System.out.println("GENERATED SUCCESS!!!");
    }
    
    /**
     * Ghi toàn bộ nội dung sang file text
     */
    public void write() {
        
        try {
            
            // Tạo đối tượng Stripper
            stripper = new PDFTextStripper();
            stripper.setStartPage(1);
            stripper.setEndPage(pdf.getNumberOfPages());
            String content = stripper.getText(pdf);
            try (BufferedWriter bw = 
                    new BufferedWriter(
                            new OutputStreamWriter(
                                    new FileOutputStream(wfile.getPath() + "\\" + foldername + ".txt"), "UTF-8"))) {
                // stripper.writeText(pdf, bw); ~NOT WORKING~
                bw.write(content);
            }
        } catch (IOException ex) {
            
            System.out.println("LOAD FILE THẤT BẠI");
        }
    }
    
    /**
     * Ghi toàn bộ nội dung sang file text của các trang cụ thể
     * @param start
     * @param end 
     */
    public void writePageRange(int start, int end) {
        
        try {
            
            // Load file pdf
            pdf = PDDocument.load(rfile);
            
            // Tạo đối tượng Stripper
            stripper = new PDFTextStripper();

            // Thiết lập các trang cần đọc
            stripper.setStartPage(start);
            stripper.setEndPage(end);
            
            try (BufferedWriter bw = 
                    new BufferedWriter(
                            new OutputStreamWriter(
                                    new FileOutputStream(wfile.getPath() + "\\" + foldername + "-page-" + start + "-to-" + end + ".txt")))) {
                stripper.writeText(pdf, bw);
            }
            
        } catch (IOException ex) {
            
            System.out.println("LOAD FILE THẤT BẠI");
        }
    }
    
    /**
     * Ghi file ảnh ra thư mục
     */
    public void writePictures() {
        PDPageTree list = pdf.getPages();
        for (PDPage page : list) {
            PDResources pdResources = page.getResources();
            int i = 1;
            for (COSName name : pdResources.getXObjectNames()) {
                PDXObject o;
                try {
                    o = pdResources.getXObject(name);
                    
                    if (o instanceof PDImageXObject) {
                        PDImageXObject image = (PDImageXObject)o;
                        String picfilename = pfile.getPath() + "\\img-" + i + ".png";
                        ImageIO.write(image.getImage(), "png", new File(picfilename));
                        i++;
                    }
                } catch (IOException ex) {
                    System.out.println("LỖI PDXObject");
                }
            }
        }
        System.out.println("LƯU ẢNH THÀNH CÔNG");
    }
    
    /**
     * Open Directory contains documents
     */
    public void open() {
        File output = new File(PATH);
        Desktop desktop = Desktop.getDesktop();
        try {
            desktop.open(output);
        } catch (IOException ex) {
            System.out.println("KHÔNG MỞ ĐƯỢC THƯ MỤC");
        }
    }
}
