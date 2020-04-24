/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package reader;

import java.awt.Desktop;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Iterator;
import java.util.List;
import java.util.UUID;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.PictureData;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author doduy
 */
public class ReadExcelFile {
    // Path for file reading
    private static final String PATH = System.getProperty("user.dir") + "\\output\\";
    // filename
    private String filename;
    // file content
    private static File rfile;
    // file pictures
    private static File pfile;
    // file document
    private static File document;
    // Create Input Stream
    private static FileInputStream excelFile = null;
    // Wordbook for excel (for open file)
    private static Workbook workbook;
    // Sheets in the workbook
    private static Sheet datatypeSheet;
    // List of Pictures
    private static List<PictureData> pictures;
   
    /**
     * Constructor for creating: read file, picture file & create folder
     * @param filename 
     */
    public ReadExcelFile(String path) {
        
        this.filename = path.substring(path.lastIndexOf('\\') + 1, path.length());
        String foldername = filename.substring(0, filename.lastIndexOf('.'));
        
        // Tạo đối tượng ghi/đọc file nội dung Workbook
        this.rfile = new File(path);
        this.document = new File(PATH + foldername + "\\documents\\"); 
        this.pfile = new File(PATH + foldername + "\\images\\");
        
        // Tạo thư mục lưu các file ảnh
        if(!pfile.exists()) {
            pfile.mkdirs();
        }
        // Tạo thư mục lưu các sheet của workbook
        if(!document.exists()) {
            document.mkdirs();
        }
        
        readWorkbook();
        System.out.println("GENERATED SUCCESS!!!");
    }
    
    /**
     * Return Workbook Object
     * @param inputStream
     * @param excelFilePath
     * @return
     * @throws IOException 
     */
    private static void readWorkbook() {
        
        try {
            excelFile = new FileInputStream(rfile);
        
            if (getFileExtension().equals(".xlsx")) {
                workbook = new XSSFWorkbook(excelFile);
            } else if (getFileExtension().equals(".xls")) {
                workbook = new HSSFWorkbook(excelFile);
            } else {
                throw new IllegalArgumentException("The specified file is not Excel file");
            }
        } catch(IOException ioe) {
            
            System.out.println("ĐỌC FILE THẤT BẠI");
        }   
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
     * Tiến hành đọc nội dung file excel
     * Ghi nội dung từng Sheet trong workbook ra file txt
     * In ra tên của từng Sheet
     */
    public void write() {
        
        if(workbook!=null) {
            String sheetName = "";
            for(int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                
                sheetName = workbook.getSheetName(sheetIdx);
                
                try (FileOutputStream out = new FileOutputStream(document.getPath() + "\\" + sheetName + ".txt")) {
                    //System.out.println("Sheet Name: " + workbook.getSheetName(sheetIdx));
                    datatypeSheet =  workbook.getSheetAt(sheetIdx);

                    // Get each row in sheet
                    Iterator<Row> iterator = datatypeSheet.iterator();

                    // Write header to file
                    // Row header = iterator.next();
                    
                    while(iterator.hasNext()) {
                        Row currentRow = iterator.next();

                        // Iterator all cells/row
                        Iterator<Cell> cellIterator = currentRow.cellIterator();

                        while(cellIterator.hasNext()) {
                            Cell currentCell = cellIterator.next();

                            switch(currentCell.getCellType()) {
                                case STRING:
                                    // System.out.print(currentCell.getStringCellValue() + "\t\t");
                                    out.write((currentCell.getStringCellValue() + "\t\t").getBytes());
                                    break;
                                case NUMERIC:
                                    // System.out.print(currentCell.getNumericCellValue() + "\t\t");
                                    out.write((currentCell.getNumericCellValue() + "\t\t").getBytes());
                                    break;
                                case BOOLEAN:
                                    // System.out.print(currentCell.getBooleanCellValue() + "\t\t");
                                    out.write((currentCell.getBooleanCellValue() + "\t\t").getBytes());
                                    break;
                                default:
                                    break;
                            }
                        }

                        out.write(("\n").getBytes());
                    }
                    System.out.println("GHI FILE HOÀN TẤT");
                } catch(IOException fnfe) {
                    System.out.println("Ghi file không thành công");
                } finally {
                    ;
                }
            }   
        }
    }
    
    /**
     * In file ra màn hình console
     */
    public void print() {
        
        if(workbook!=null) {
            
            for(int sheetIdx = 0; sheetIdx < workbook.getNumberOfSheets(); sheetIdx++) {
                
                String sheetName = workbook.getSheetName(sheetIdx);
                
                System.out.println("Sheet Name: " + workbook.getSheetName(sheetIdx));
                datatypeSheet =  workbook.getSheetAt(sheetIdx);

                // Get each row in sheet
                Iterator<Row> iterator = datatypeSheet.iterator();

                while(iterator.hasNext()) {
                    Row currentRow = iterator.next();

                    // Iterator all cells/row
                    Iterator<Cell> cellIterator = currentRow.cellIterator();

                    while(cellIterator.hasNext()) {
                        Cell currentCell = cellIterator.next();

                        switch(currentCell.getCellType()) {
                            case STRING:

                                System.out.print(currentCell.getStringCellValue() + "\t\t");
                                break;
                            case NUMERIC:

                                System.out.print(currentCell.getNumericCellValue() + "\t\t");
                                break;
                            case BOOLEAN:

                                System.out.print(currentCell.getBooleanCellValue() + "\t\t");
                                break;
                            default:
                                break;
                        }
                    }

                    System.out.println();
                }
            }   
        }
    }
    
    /**
     * Lưu tất cả ảnh có trong workbook
     */
    public void writePictures() {

        // Đọc file hình
        pictures = (List<PictureData>) workbook.getAllPictures();

        // System.out.println("Đã tìm được " + pictures.size() + " ảnh.");
            
        for (PictureData pict : pictures) {

            String ext = pict.suggestFileExtension();

            byte[] data = pict.getData();

            // Ghi File ảnh ra định dạng: png hoặc jpg
            try (FileOutputStream out = new FileOutputStream(pfile.getPath() + "\\img-" + UUID.randomUUID() + "." + ext)){

                // Ghi file
                out.write(data);

            } catch (IOException ex) {
                System.out.println("GHI FILE ẢNH THẤT BẠI.");
            } finally {
                ;
            }
        }
        System.out.println("ĐÃ LƯU TẤT CẢ ẢNH TRONG WORKBOOK!!!");
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
