import org.apache.poi.ss.usermodel.Cell;

/**
 * Created by Arif on Окт., 2018
 */

public class Main {
    //var20 have to exist, if var20.xlsx don't exist then throw IOException
    //To correct work this program the file have to close in your system!!!
    private static final String filePath = "var20.xlsx";

    public static void main(String[] args) {
        IOCell ioCell = new IOCell(filePath);

        Cell x = ioCell.getCell(0, 1, 0);
        Cell y = ioCell.getCell(0,  1, 1);
        System.out.println("first number: " + x.toString());
        System.out.println("second number: " + y.toString());
        //Write x * y
        ioCell.setCell(4, 0, x.getNumericCellValue() * y.getNumericCellValue());
        //Write x + y
        ioCell.setCell(4, 1, x.getNumericCellValue() + y.getNumericCellValue());
        System.out.println("Interactions is complete successfully");
    }
}
