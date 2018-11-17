**Интеграция электронных таблиц MS Excel и Java.**
---

### Описание:

В современном мире очень много случаев, при которых необходимо интегрировать MS
Excel с Java. Например, при разработке Enterprise-приложения в некой финансовой
сфере, вам необходимо предоставить счет для заинтересованных лиц, а проще всего
выставлять счет на MS Excel.

### Обзор существующих API MS Excel для Java:
Рассмотрим основные API:
* Docx4j - это  API с открытым исходным кодом, для создания и манипулирования документами формата Microsoft Open XML, к которым отросятся Word docx, Powerpoint pptx, Excel xlsx файлы. Он очень похож на Microsoft OpenXML SDK, но реализован на языке Java. Docx4j использует JAXB архитектуру для создания представления объекта в памяти. Docx4j акцентирует свое внимание на всесторонней поддержке заявленного формата, но от пользователя данного API требуется знание и понимание технологии JAXB и структуры Open XML.

* Apache POI - это набор API с открытым исходным кодом, который предлагает определенные функции для чтения и записи различных документов, базирующихся на Office Open XML стандартах (OOXML) и Microsoft OLE2 форматe документов (OLE2). OLE2 файлы включают большинство Microsoft Office форматов, таких как doc, xls, ppt. Office Open XML формат это новый стандарт базирующийся на XML разметке, и используется в файлах Microsoft office 2007 и старше.

* Aspose for Java - набор платных Java APIs, которые помогают разработчикам в работе с популярными форматами бизнес файлов, такими как документы Microsoft Word, таблицы Microsoft Excel, презентации Microsoft PowerPoint, PDF файлы Adobe Acrobat, emails, изображения, штрих-коды и оптические распознавания символов.

Каждое API проектируется для того, чтобы выполнять широкий спектр создания документов, различные манипуляции и преобразования быстро и легко, экономя время и позволяя разработчикам успешно программировать. Ни один API с открытым исходным кодом не имеет одной и той же комплексной поддержки функций.

Все Aspose’s APIs используют простую объектную модель документа, а одно API предназначено для работы с набором связанных форматов. Aspose’s Microsoft Office APIs, Aspose.Cells, Aspose.Words, Aspose.Slides, Aspose.Email, и Aspose.Tasks легки в работе, эффективны, надежны и независимы от других библиотек.

Преимуществом APIs с открытым исходным кодом является то, что они бесплатны и каждый может настроить их под свои задачи и цели. Это очень удобно, если у пользователя есть достаточно времени и ресурсов. Однако данные APIs не всегда имеют поддержку или документацию, и поддерживают небольшое количество функций и вариантов. Этот недостаток стоит разработчикам времени, и сокращает надежность их приложений. К преимуществам проприетарных (коммерческих) API можно отнести комплексную поддержку функционала с подробной документацией, регулярное обновление, гарантию отсутствия ошибок и обратную связь с разработчиками APIs.

В данной программе будем использовать Apache POI

### Ссылки на полезные ресурсы
* https://habr.com/post/56817/
* https://poi.apache.org/apidocs/index.html --- официальная документация
* http://java-online.ru/java-excel.xhtml


### Задание:

В данной работе вы должны реализовать следующее:

1.  Чтение с ячейки MS Excel в Java

2.  Запись с Java в MS Excel

### Инструкция
   
* Для обращения к MS Excel версии до 2003 включительно года с Java используется класс ```HSSFWorkbook```
* Для обращения к MS Excel версии 2007 и позднее с Java используется класс ```XSSFWorkbook```
* При операциях ***Обновление*** или ***Запись*** необходимо, чтобы MS Excel был закрыт.

##### Чтение ячейки с MS Excel 
Чтобы считать данные с ```xlsx``` необходимо исполнить следующие шаги:
```java
    //filePath - это путь до MS Excel
    Workbook book = new XSSFWorkbook(new FileInputStream(filePath);
    //считывается лист по индексу sheet_index. sheet_index начинается с 0
    Sheet sheet = book.getSheetAt(sheet_index);
    //считывается row по индексу row_index. row_index начинается с 0
    Row row = sheet.getRow(row_index);
    //считывается cell по индексу cell_index. cell_index начинается с 0
    Cell cell = sheet.getCell(cell_index);
```

##### Запись в ячейку MS Excel
```java
    Workbook book = new XSSFWorkbook();
    //name - имя листа
    Sheet sheet = book.createSheet(name);
    Row row = sheet.createRow(i);
    Cell cell = row.createCell(j);
    FileInputStream fileOut = new FileInputStream(filePath);
    book.write(fileOut);
    fileOut.close();
```

##### Обновление ячейки в существующем листе MS Excel
```java
    Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
    Sheet sheet = workbook.getSheetAt(i);
    Row row = sheet.getRow(j);
    Cell cell = row.getCell(k);
    cell.setCellValue(value);
```

### Выполнение:

1.  Создать проект на java с помощью maven.

2.  Прописать следующие зависимости в pom.xml:
    

3.  Создать Excel файл в корневой папке проекта.

4.  Записать в A1 и A2 любые целые числа.

5.  В папке src/main/java создать класс IOCell

    1.  Создать поле 
    ```java
    private File filePath
    ```

    2.  Создать конструктор 
    ```java
    IOCell(String filePath) { this.filePath = new File(filePath)}
    ```

    3.  Создать метод для чтения c Excel в Java
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
    4.  Создать метод для записи с Java в Excel 
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

6.  В папке src/main/java создать класс Main
        
    1. Создать поле  
    ```java
        private static final String filePath = "NAME_OF_EXCEL_FILE";
    ```
    1.  Создать метод
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
7. Запускаем приложение и смотрим в консоль.