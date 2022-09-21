import java.io.*;
import java.util.*;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Main {
    public static void main(String[] args) throws Exception {
        ArrayList<String> valuesColumn= new  ArrayList<>();
        int[][] numbers=readExcel(valuesColumn);
        if(numbers!=null)
        {
            Scanner scan = new Scanner (System.in);
            System.out.print("Введите как вы хотите вывести результат:\n1 - в excel файл\n2 - в консоль\nВаш выбор: ");
            String resultOutput= scan.nextLine();
            if(resultOutput.equals("1")||resultOutput.equals("2"))
            {
                if(resultOutput.equals("1"))
                {
                    System.out.print("Введите путь для нового excel файла: ");
                    String path= scan.nextLine();
                    Workbook workbookWrite = new XSSFWorkbook();
                    Sheet newSheet = workbookWrite.createSheet("Результат");
                    dataProcessing(numbers,numbers.length,valuesColumn.size(),valuesColumn,path,workbookWrite,newSheet,resultOutput);
                    System.out.println("Файл создан");
                }else{
                    dataProcessing(numbers,numbers.length,valuesColumn.size(),valuesColumn,null,null,null,resultOutput);
                }
            }else{
                System.out.println("Такого выбора нет!!!");
            }
        }
        valuesColumn.clear();
    }
    public static int[][] readExcel(ArrayList<String> valuesColumn)///считывае данные из excel файла: значения(способы использования каждой из колоной) и двумерный массив чисел
    {
        try {
            Scanner scan = new Scanner (System.in);
            System.out.print("Введите путь к excel файлу: ");
            String path = scan.nextLine();
            FileInputStream file = new FileInputStream(new File(path));
            XSSFWorkbook workbookRead = new XSSFWorkbook(file);
            System.out.print("Введите номер страницы: ");
            int sheetNumber = scan.nextInt();
            XSSFSheet sheet = workbookRead.getSheetAt(sheetNumber);
            Iterator<Row> rowIterator = sheet.iterator();
            //////считывание данных//////
            int[][] numbers= new int[(sheet.getLastRowNum()+1-1)][(sheet.getRow(0).getPhysicalNumberOfCells())];
            while (rowIterator.hasNext()) {
                Row row = rowIterator.next();
                Iterator<Cell> cellIterator = row.cellIterator();
                while (cellIterator.hasNext()) {
                    Cell cell = cellIterator.next();
                    switch (cell.getCellType()){
                        case NUMERIC:
                            if(row.getRowNum()>0){
                                numbers[(row.getRowNum()-1)][cell.getColumnIndex()]= (int)cell.getNumericCellValue();
                            }
                            break;
                        case STRING:
                            if (row.getRowNum()==0) {
                                valuesColumn.add(cell.getStringCellValue());
                            }
                            break;
                    }
                }
            }
            return numbers;
        } catch (Exception e) {
            System.out.println("Что-то пошло не так при чтении файла!!!");
            return null;
        }
    }
    public static void dataProcessing(int[][] data, int row,int column,ArrayList<String> valuesColumn,String path,Workbook workbookWrite,Sheet newSheet,String resultOutput)///обработка значений и массива с числами
    {
        ArrayList<String> result= new  ArrayList<>();
        for (int i = 0; i <valuesColumn.size(); i++) {
            if(!valuesColumn.get(i).equals("-")){
                result.add(valuesColumn.get(i));
            }
        }
        int resultRow=0;
        ///вывод способа использования каждой колонки выбранным способом (на консоль или в файл)
        if(resultOutput.equals("1"))
        {
            printExcel(result,resultRow,path,workbookWrite,newSheet);
        }else{
            System.out.println("Результат:");
            printConsole(result);
        }
        resultRow++;
        result.clear();
        Integer checkedRows=0;
        Integer numberCoincidences = 0;
        int group= 0;
        int curRow=0;
        int curColumnResult=0;
        int sum=0;
        int max=0;
        int min=0;
        String konk="";
        while(!checkedRows.equals(row))///выполнять пока не будут проверены все строки в матрице
        {
            group=0;///находим количество строк для группы по критерию (в нашем случае по "NUM")
            for (int i = 0; i < column; i++) {
                if (valuesColumn.get(i).equals("NUM")) {
                    numberCoincidences = 0;
                    curRow = (int) checkedRows;
                    while ((curRow < row) && data[checkedRows][i] == data[curRow][i]) {
                        curRow++;
                        numberCoincidences++;
                    }
                    if (group == 0) {
                        group = numberCoincidences;
                    } else if (numberCoincidences < group) {
                        group = numberCoincidences;
                    }
                }
            }
                curColumnResult=0;
                while (curColumnResult<valuesColumn.size())///добавляем в временный список все колонки
                {
                    if(!valuesColumn.get(curColumnResult).equals("-"))///выполнять, если эта колонка не с критерием "-"
                    {
                        switch (valuesColumn.get(curColumnResult)){
                            case "NUM":///просто записываем 1ую строчку данного столбца (они все одинаковы) группы
                                result.add(Integer.toString(data[checkedRows][curColumnResult]));
                                break;
                            case "SUM":///суммируем все строчки группы данного столбца
                                sum=0;
                                for (int j = checkedRows; j < checkedRows+group; j++) {
                                    sum+=(int)data[j][curColumnResult];
                                }
                                result.add(Integer.toString(sum));
                                break;
                            case "MIN":///находим минимальное значение среди всех строчек группы данного столбца
                                min=data[checkedRows][curColumnResult];
                                for (int j = checkedRows+1; j < checkedRows+group; j++) {
                                    if(min>data[j][curColumnResult]){
                                        min=data[j][curColumnResult];
                                    }
                                }
                                result.add(Integer.toString(min));
                                break;
                            case "MAX":///находим максимальное значение среди всех строчек группы данного столбца
                                max=data[checkedRows][curColumnResult];
                                for (int j = checkedRows+1; j < checkedRows+group; j++) {
                                    if(max<data[j][curColumnResult]){
                                        max=data[j][curColumnResult];
                                    }
                                }
                                result.add(Integer.toString(max));
                                break;
                            case "CONC":///производим конкатенацию всех строчек группы данного столбца
                                konk="";
                                for (int j = checkedRows; j < checkedRows+group; j++) {
                                    konk+=Integer.toString(data[j][curColumnResult]);
                                }
                                result.add(konk);
                                break;
                        }
                    }
                    curColumnResult++;
                }
            if(resultOutput.equals("1"))///вывод группы выбранным способом (на консоль или в файл)
            {
                printExcel(result,resultRow,path,workbookWrite,newSheet);
            }else{
                printConsole(result);
            }
            resultRow++;
            result.clear();
            checkedRows+=group;///смещаем "указатель" проверенных строк матрицы на количество строк в данной группе
        }
    }
    public static void printExcel(ArrayList<String> data,int indexRow,String path,Workbook workbookWrite,Sheet newSheet)///запись полученного списка в новый excel файл
    {
        Row row;
        row = newSheet.createRow(indexRow);
        for (int i = 0; i < data.size(); i++) {
            row.createCell(i).setCellValue(data.get(i));
        }
        try {
            FileOutputStream fileOut = new FileOutputStream(path);
            workbookWrite.write(fileOut);
            fileOut.close();
        }catch (Exception e){
            System.out.println("Что-то пошло не так при записи в файл!!!");
        }
    }
    public static void printConsole(ArrayList<String> data)///вывод полученного списка на консоль
    {
        for (int i = 0; i < data.size(); i++) {
            System.out.printf("%-10s",data.get(i));
        }
        System.out.println();
    }
}