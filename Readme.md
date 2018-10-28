# Интеграция электронных таблиц MS Excel и Java.

Данная программа демонстрирует использование Apache POI, которая предоставляет API MS Excel для Java.
* Для обращения к MS Excel версии до 2003 включительно года с Java используется класс ```HSSFWorkbook```
* Для обращения к MS Excel версии 2007 и позднее с Java используется класс ```XSSFWorkbook```
* При операциях ***Обновление*** или ***Запись*** необходимо, чтобы MS Excel был закрыт.

## Чтение ячейки с MS Excel 
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

## Запись в ячейку MS Excel
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

## Обновление ячейки в существующем листе MS Excel
```java
    Workbook workbook = new XSSFWorkbook(new FileInputStream(filePath));
    Sheet sheet = workbook.getSheetAt(i);
    Row row = sheet.getRow(j);
    Cell cell = row.getCell(k);
    cell.setCellValue(value);
```
