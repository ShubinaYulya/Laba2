package laba2;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.file.FileSystems;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import javafx.scene.control.Cell;
import org.apache.commons.math3.stat.descriptive.DescriptiveStatistics;
import org.apache.poi.sl.draw.geom.Path;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;


public class ExcelManipulator {//эксель-манипулятор
public ExcelManipulator() {};//конструктор по умолчанию
HashMap<String, double[]> MyExport = new HashMap();//хеш-карта со строкой и массивом чисел (нужно для экспорта)
public void export() throws FileNotFoundException, IOException {//функция экспорта из файла
        
java.nio.file.Path file_path = FileSystems.getDefault().getPath("ДЗ2.xlsx");//открываем файл
XSSFWorkbook MyBook = new XSSFWorkbook(new FileInputStream(file_path.toString()));//создаем книгу
XSSFSheet MySheet = MyBook.getSheet("Вариант 10");//находим лист под названием вариант 10
int rowCount = MySheet.getPhysicalNumberOfRows();//получаем число строк
XSSFRow headers = MySheet.getRow(0);//получаем список заголовков в первой строке
for (int i=0; i<headers.getPhysicalNumberOfCells() ; i++) {//обходим каждый столбец
XSSFCell header = headers.getCell(i);//заголовок из ячейки
String ColName = header.getStringCellValue();//название колонки
double[] values = new double[rowCount-1];//значения колонки (числа)
int k = 0;
for (int j=1; j<rowCount; j++) {//обходим всю колонку вниз
values[k] = MySheet.getRow(j).getCell(i).getNumericCellValue();//добавляем числа из колонки
k++;
}
MyExport.put(ColName, values);//добавляем в хеш-карту соотвевтсвие названия колонки и массива чисел
}
//System.out.println("Импорт выполнен");//импорт выполнен
}

public void result() throws FileNotFoundException, IOException {//выводим результат (он же экспорт)
XSSFWorkbook MyBook = new XSSFWorkbook();//новая эксель-книга
XSSFSheet MySheet = MyBook.createSheet("Вариант 10");//листок вариант 10
Row row1 = MySheet.createRow(0);//созаем три строки
Row row2 = MySheet.createRow(1);
Row row3 = MySheet.createRow(2);
org.apache.poi.ss.usermodel.Cell mx = row2.createCell(0);
mx.setCellValue("max");//максимум 
org.apache.poi.ss.usermodel.Cell mn = row3.createCell(0);
mn.setCellValue("min");//минимум
int i=1;
for (Map.Entry<String, double[]> pair:MyExport.entrySet()) {//обходим нашу эксель-страницу(получена в импорте)
DescriptiveStatistics descriptiveStatistics = new DescriptiveStatistics();//объект, реализующий статистические вычисления
double[] vals = pair.getValue();
for (double v : vals) {
descriptiveStatistics.addValue(v);//добавляем все данные из столбца
}
double max = descriptiveStatistics.getMax();//получаем минимум и максимум
double min = descriptiveStatistics.getMin();

org.apache.poi.ss.usermodel.Cell header = row1.createCell(i);
header.setCellValue(pair.getKey());//вставляем в ячейки полученные значения
org.apache.poi.ss.usermodel.Cell maxcell = row2.createCell(i);
maxcell.setCellValue(max);
org.apache.poi.ss.usermodel.Cell mincell = row3.createCell(i);
mincell.setCellValue(min);
i++;
}
        
try {
MyBook.write(new FileOutputStream("Расчёты.xlsx"));//сохраняем эксель-файл
} catch (IOException ex) {
Logger.getLogger(ExcelManipulator.class.getName()).log(Level.SEVERE, null, ex);//если получили исключение
}
MyBook.close();
//System.out.println("Экспорт выполнен");
}
}
