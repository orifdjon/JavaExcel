**���������� ����������� ������ MS Excel � Java.**
---

### ��������:

� ����������� ���� ����� ����� �������, ��� ������� ���������� ������������� MS
Excel � Java. ��������, ��� ���������� Enterprise-���������� � ����� ����������
�����, ��� ���������� ������������ ���� ��� ���������������� ���, � ����� �����
���������� ���� �� MS Excel.

### ����� ������������ API MS Excel ��� Java:
���������� �������� API:
* Docx4j - ���  API � �������� �������� �����, ��� �������� � ��������������� ����������� ������� Microsoft Open XML, � ������� ��������� Word docx, Powerpoint pptx, Excel xlsx �����. �� ����� ����� �� Microsoft OpenXML SDK, �� ���������� �� ����� Java. Docx4j ���������� JAXB ����������� ��� �������� ������������� ������� � ������. Docx4j ����������� ���� �������� �� ������������ ��������� ����������� �������, �� �� ������������ ������� API ��������� ������ � ��������� ���������� JAXB � ��������� Open XML.

* Apache POI - ��� ����� API � �������� �������� �����, ������� ���������� ������������ ������� ��� ������ � ������ ��������� ����������, ������������ �� Office Open XML ���������� (OOXML) � Microsoft OLE2 ������e ���������� (OLE2). OLE2 ����� �������� ����������� Microsoft Office ��������, ����� ��� doc, xls, ppt. Office Open XML ������ ��� ����� �������� ������������ �� XML ��������, � ������������ � ������ Microsoft office 2007 � ������.

* Aspose for Java - ����� ������� Java APIs, ������� �������� ������������� � ������ � ����������� ��������� ������ ������, ������ ��� ��������� Microsoft Word, ������� Microsoft Excel, ����������� Microsoft PowerPoint, PDF ����� Adobe Acrobat, emails, �����������, �����-���� � ���������� ������������� ��������.

������ API ������������� ��� ����, ����� ��������� ������� ������ �������� ����������, ��������� ����������� � �������������� ������ � �����, ������� ����� � �������� ������������� ������� ���������������. �� ���� API � �������� �������� ����� �� ����� ����� � ��� �� ����������� ��������� �������.

��� Aspose�s APIs ���������� ������� ��������� ������ ���������, � ���� API ������������� ��� ������ � ������� ��������� ��������. Aspose�s Microsoft Office APIs, Aspose.Cells, Aspose.Words, Aspose.Slides, Aspose.Email, � Aspose.Tasks ����� � ������, ����������, ������� � ���������� �� ������ ���������.

������������� APIs � �������� �������� ����� �������� ��, ��� ��� ��������� � ������ ����� ��������� �� ��� ���� ������ � ����. ��� ����� ������, ���� � ������������ ���� ���������� ������� � ��������. ������ ������ APIs �� ������ ����� ��������� ��� ������������, � ������������ ��������� ���������� ������� � ���������. ���� ���������� ����� ������������� �������, � ��������� ���������� �� ����������. � ������������� ������������� (������������) API ����� ������� ����������� ��������� ����������� � ��������� �������������, ���������� ����������, �������� ���������� ������ � �������� ����� � �������������� APIs.

� ������ ��������� ����� ������������ Apache POI

### ������ �� �������� �������
* https://habr.com/post/56817/
* https://poi.apache.org/apidocs/index.html --- ����������� ������������
* http://java-online.ru/java-excel.xhtml


### �������:

� ������ ������ �� ������ ����������� ���������:

1.  ������ � ������ MS Excel � Java

2.  ������ � Java � MS Excel

### ����������
   
* ��� ��������� � MS Excel ������ �� 2003 ������������ ���� � Java ������������ ����� ```HSSFWorkbook```
* ��� ��������� � MS Excel ������ 2007 � ������� � Java ������������ ����� ```XSSFWorkbook```
* ��� ��������� ***����������*** ��� ***������*** ����������, ����� MS Excel ��� ������.

##### ������ ������ � MS Excel 
����� ������� ������ � ```xlsx``` ���������� ��������� ��������� ����:
```java
    //filePath - ��� ���� �� MS Excel
    Workbook book = new XSSFWorkbook(new FileInputStream(filePath);
    //����������� ���� �� ������� sheet_index. sheet_index ���������� � 0
    Sheet sheet = book.getSheetAt(sheet_index);
    //����������� row �� ������� row_index. row_index ���������� � 0
    Row row = sheet.getRow(row_index);
    //����������� cell �� ������� cell_index. cell_index ���������� � 0
    Cell cell = sheet.getCell(cell_index);
```

##### ������ � ������ MS Excel
```java
    Workbook book = new XSSFWorkbook();
    //name - ��� �����
    Sheet sheet = book.createSheet(name);
    Row row = sheet.createRow(i);
    Cell cell = row.createCell(j);
    FileInputStream fileOut = new FileInputStream(filePath);
    book.write(fileOut);
    fileOut.close();
```

##### ���������� ������ � ������������ ����� MS Excel
```java
    Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
    Sheet sheet = workbook.getSheetAt(i);
    Row row = sheet.getRow(j);
    Cell cell = row.getCell(k);
    cell.setCellValue(value);
```

### ����������:

1.  ������� ������ �� java � ������� maven.

2.  ��������� ��������� ����������� � pom.xml:
    

3.  ������� Excel ���� � �������� ����� �������.

4.  �������� � A1 � A2 ����� ����� �����.

5.  � ����� src/main/java ������� ����� IOCell

    1.  ������� ���� 
    ```java
    private File filePath
    ```

    2.  ������� ����������� 
    ```java
    IOCell(String filePath) { this.filePath = new File(filePath)}
    ```

    3.  ������� ����� ��� ������ c Excel � Java
    ```java
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
    ```
    4.  ������� ����� ��� ������ � Java � Excel 
    ```
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
    ```

6.  � ����� src/main/java ������� ����� Main
        
    1. ������� ����  
    ```java
        private static final String filePath = "NAME_OF_EXCEL_FILE";
    ```
    1.  ������� �����
    ```java
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
    ```
7. ��������� ���������� � ������� � �������.