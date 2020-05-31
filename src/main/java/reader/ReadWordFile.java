/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package reader;

import java.awt.Desktop;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FilenameFilter;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStreamWriter;
import java.util.List;
import java.util.UUID;
import org.apache.poi.hwpf.HWPFDocument;
import org.apache.poi.hwpf.model.PicturesTable;
import org.apache.poi.ooxml.extractor.POIXMLTextExtractor;
import org.apache.poi.hwpf.usermodel.Picture;
import org.apache.poi.xwpf.extractor.XWPFWordExtractor;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFPictureData;
/**
 *
 * @author doduy
 */
public class ReadWordFile {
    // Path for file reading
    private static final String PATH = System.getProperty("user.dir") + "\\output\\";
    // Path for picture reading store
    //private static final String IPATH = System.getProperty("user.dir") + "\\demofiles\\images\\";
    // filename
    private String filename;
    // file content
    private static File rfile;
    // file pictures
    private static File pfile;
    // file document
    private static File wfile;
    // Container of document text
    private static String content;
    // Create Input Stream
    private static InputStream inputStream = null;
    // Use for reading DOC file 
    private static HWPFDocument doc = null;
    // Use for reading DOCX file
    private static XWPFDocument docx = null;
    // Container of Document
    private static POIXMLTextExtractor extractor = null; 
    // List of pictures
    private static List<Picture> pictures;
    // List of pictures
    private static List<XWPFPictureData> xpictures;
    // name of folder
    private static String foldername;
    
    /**
     * Default Constructor
     */
    public ReadWordFile() {}
    
    /**
     * Constructor for creating: read file, picture file & create folder
     * @param filename 
     */
    public ReadWordFile(String path) {
        
        this.filename = path.substring(path.lastIndexOf('\\') + 1, path.length());
        this.filename = filename;
        foldername = filename.substring(0, filename.lastIndexOf('.'));
        
        // Tạo đối tượng ghi/đọc file nội dung Workbook
        this.rfile = new File(path);
        this.wfile = new File(PATH + foldername + "\\documents\\"); 
        this.pfile = new File(PATH + foldername + "\\images\\");
        
        // Tạo thư mục lưu các file ảnh
        if(!pfile.exists()) {
            pfile.mkdirs();
        }
        // Tạo thư mục lưu văn bản
        if(!wfile.exists()) {
            wfile.mkdirs();
        }
        
        readWordDoc();
        System.out.println("GENERATED SUCCESS!!!");
    }
    
    /**
     * Return extension of a file object
     * @param file
     * @return 
     */
    private static String getFileExtension() {
        String extension = "";
 
        try {
            if (rfile != null && rfile.exists()) {
                String name = rfile.getName();
                extension = name.substring(name.lastIndexOf("."));
            }
        } catch (Exception e) {
            extension = "";
        }
 
        return extension;
    }
    
    /**
     * Return Workbook Object
     * @param inputStream
     * @param excelFilePath
     * @return
     * @throws IOException 
     */
    private static void readWordDoc() {
        
        if(rfile != null) {
            
            try {
                
                inputStream = new FileInputStream(rfile);
                if (getFileExtension().equals(".docx")) {
                    
                    docx = new XWPFDocument(inputStream);
                } else if (getFileExtension().equals(".doc")) {
                    
                    doc = new HWPFDocument(inputStream);
                } else {
                    
                    throw new IllegalArgumentException("KHÔNG TÌM THẤY ĐỊNH DẠNG FILE DOC HOẶC DOCX");
                }
            } catch (FileNotFoundException fnfe) {
                
                System.out.println(fnfe.toString() + ": LỖI KHÔNG TÌM THẤY FILE!!!");
            } catch (IOException ioe) {
                
                System.out.println(ioe.toString() + ": LỖI ĐỌC FILE");
            } finally {
                
                try {
                    inputStream.close();
                } catch (IOException ioe) {
                    
                    System.out.println(ioe.toString() + ": ĐÓNG FILE THẤT BẠI");
                }
            }
        } else {
            System.out.println("Cannot Read this file");
        }
    }
    
    /**
     * In ra màn hình console
     */
    public void print() {        
        
        if (doc!=null) {
            
            content = doc.getDocumentText();
            System.out.println(content);
        } else if (docx!=null) {

            extractor = new XWPFWordExtractor(docx);
            content = extractor.getText();
            System.out.println(content);
        } else {
            System.out.println("CHƯA ĐỌC FILE!!!");
        }
    }
    
    /**
     * Ghi nội dung word ra file txt
     * Ghi các thông tin cơ bản của file word
     * Cắt bỏ khoảng trống, dòng trống
     * Loại bỏ numbering và bullet
     */
    public void write() {        
        
        if (doc!=null) {
            content = doc.getDocumentText();
        } else if (docx!=null) {
            extractor = new XWPFWordExtractor(docx);
            content = extractor.getText();
        } else {
            System.out.println("CHƯA ĐỌC FILE!!!");
        }
        
        try (BufferedWriter bw = 
            new BufferedWriter(
                    new OutputStreamWriter(
                            new FileOutputStream(wfile.getPath() + "\\" + foldername + ".txt"), "UTF-8"))) {
            bw.write(content);
        } catch (IOException ex) {
            System.out.println("LOAD FILE THẤT BẠI");
        }
    }
    
    /**
     * Trích xuất ảnh trong word ra file riêng
     */
    public void writePictures() {
        if (doc!=null) {
            
            // Lấy tất cả hình trong WORD
            PicturesTable picturesTable = doc.getPicturesTable();
            pictures = picturesTable.getAllPictures();
            
            // Ghi từng file word ra theo định dạng ảnh
            for (Picture picture : pictures) {
                String ext = picture.suggestFileExtension();
                byte[] data = picture.getContent();

                try (FileOutputStream out = new FileOutputStream(pfile.getPath() + "\\" + UUID.randomUUID() + "." + ext)){

                    out.write(data);
                } catch (IOException fnfe) {

                    System.out.println("GHI FILE ẢNH THẤT BẠI!!!");
                } finally {
                    ;
                }
                //picture.writeImageContent(out);
            }
            System.out.println("LƯU ẢNH HOÀN TẤT");
        } else if (docx!=null) {

            // Document image content
            xpictures = docx.getAllPictures();
            for (XWPFPictureData picture : xpictures) {
                String ext = picture.suggestFileExtension();
                byte[] data = picture.getData();
                try (FileOutputStream out = new FileOutputStream(pfile.getPath() + "\\img-" + UUID.randomUUID() + "." + ext)){

                    out.write(data);
                } catch (IOException fnfe) {

                    System.out.println("GHI FILE ẢNH THẤT BẠI!!!");
                } finally {
                    ;
                }
            }
            System.out.println("LƯU ẢNH HOÀN TẤT");
        } else {
            System.out.println("BẠN CHƯA ĐỌC FILE!!!");
        }
    }
    
    /**
     * Return path file
     * @return 
     */
    public String getFilePath() {
        return rfile.getPath();
    }
    
    /**
     * Return picture path file
     * @return 
     */
    public String getPicturePath() {
        return pfile.getPath();
    }
    
    public String getOutputPath() {
        return PATH;
    }
    
    /**
     * Open Directory contains documents
     */
    public static void open() {
        File output = new File(PATH);
        Desktop desktop = Desktop.getDesktop();
        try {
            desktop.open(output);
        } catch (IOException ex) {
            System.out.println("KHÔNG MỞ ĐƯỢC THƯ MỤC");
        }
    }
}
