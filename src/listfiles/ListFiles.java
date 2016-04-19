/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package listfiles;

import java.io.BufferedReader;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.Iterator;
import java.util.stream.Stream;
import org.apache.commons.io.FilenameUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author Carlos Cortina
 */
public class ListFiles {
    
    private static Cell checkRowCellExists(XSSFSheet currentSheet,int rowIndex, int colIndex){
        Row currentRow = currentSheet.getRow(rowIndex);
        if( currentRow == null){
            currentRow = currentSheet.createRow(rowIndex);
        }
        //Check if cell exists
        Cell currentCell = currentRow.getCell(colIndex);
        if( currentCell == null){
            currentCell = currentRow.createCell(colIndex);
        }
        return currentCell;
    }

    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) {
        // TODO code application logic here
        BufferedReader reader = new BufferedReader(new InputStreamReader(System.in));
        String folderPath = "";
        String fileName = "DirectoryFiles.xlsx";
        
        try {
            System.out.println("Folder path :");
            folderPath = reader.readLine();
            //System.out.println("Output File Name :");
            //fileName = reader.readLine();
            
            XSSFWorkbook wb = new XSSFWorkbook();
            FileOutputStream fileOut = new FileOutputStream(folderPath+"\\"+fileName);
            XSSFSheet sheet1 = wb.createSheet("Files");
            int row = 0;
            Stream<Path> stream = Files.walk(Paths.get(folderPath));
            Iterator<Path> pathIt= stream.iterator();
            String ext = "";
            
            while(pathIt.hasNext()){
                Path filePath = pathIt.next();
                Cell cell1 = checkRowCellExists(sheet1, row, 0);
                Cell cell2 = checkRowCellExists(sheet1, row, 1);
                row++;
                ext = FilenameUtils.getExtension(filePath.getFileName().toString());
                cell1.setCellValue(filePath.getFileName().toString());
                cell2.setCellValue(ext);
                
            }
            sheet1.autoSizeColumn(0);
            sheet1.autoSizeColumn(1);
            
            wb.write(fileOut);
            fileOut.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
        System.out.println("Program Finished");
    }
    
}
