import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.*;

/**
 * Created by Arif on Окт., 2018
 */
public class IOCell {

    private File filePath;

    private IOCell() {
    }

    public IOCell(String filePath) {
        this.filePath = new File(filePath);
    }

    /**
     * Method return cell. if filePath is specified in constructor is not exist, throw exception
     *
     * @param sheet  with 0
     * @param row    with 0
     * @param column with 0
     * @return cell
     */
    public Cell getCell(int sheet, int row, int column) {
        Workbook workbook = null;
        try (FileInputStream file = new FileInputStream(filePath)) {
            workbook = new XSSFWorkbook(file);
        } catch (FileNotFoundException e) {
            System.out.println("file is not exists");
        } catch (IOException e) {
            e.printStackTrace();
        }
        return workbook.getSheetAt(sheet).getRow(row).getCell(column);
    }


    public void setCell(int row, int column, double val) {
        Workbook workbook = null;
         try (FileInputStream file = new FileInputStream(filePath)) {
             workbook = new XSSFWorkbook(file);
             Sheet sheet = workbook.getSheetAt(0);
             sheet.getRow(row).getCell(column).setCellValue(val);
         } catch (IOException e) {
             e.printStackTrace();
         }
        try (OutputStream fileOut = new FileOutputStream(filePath)) {
            workbook.write(fileOut);
        } catch (FileNotFoundException e) {
            System.out.println("file is not exist AAAA");
        } catch (IOException e) {
            e.printStackTrace();
        }

    }
}
