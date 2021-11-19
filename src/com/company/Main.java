package com.company;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Scanner;
import java.util.TreeMap;

public class Main {
    public static Integer key = 0;
    public static Scanner sn = new Scanner(System.in);
    public static XSSFWorkbook wb = new XSSFWorkbook();

    public static void main(String[] args) throws IOException {
//        FileInputStream f = new FileInputStream("E:\\javaFileHandling\\stdDetails.xlsx");
//        XSSFWorkbook workbook = new XSSFWorkbook(f);
//        XSSFSheet sheet = workbook.getSheetAt(0);
//        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
//        for(Row r : sheet){
//            for(Cell c : r){
//                switch (c.getCellType()){
//                    case Cell.CELL_TYPE_STRING :
//                        System.out.print(c.getStringCellValue()+"    ");
//                        break;
//                    case Cell.CELL_TYPE_NUMERIC:
//                        System.out.print(c.getNumericCellValue()+"   ");
//                        break;
//                }
//            }
//            System.out.println();
//        }
//
//        System.out.println();
//
//        Row r = sheet.getRow(2);
//        Cell c = r.getCell(0);
//        System.out.println(c.getStringCellValue());

        XSSFSheet sheet1 = wb.createSheet("Student Data1");
        Map<Integer,Object[]> map = new TreeMap<>();
        map.put(0,new Object[]{
                "ID","Name","Marks"
        });
        System.out.print("Number of records you want to enter into excel1 : ");
        int n = sn.nextInt();
        for(int i=0;i<n;i++){
            addRecordToMap(map);
        }
        key = 0;

        addAllRecordsToSheet(sheet1,map);






        XSSFSheet sheet2 = wb.createSheet("Student Data2");
        map = new TreeMap<>();
        map.put(0,new Object[]{
                "ID","Name","Marks"
        });
        System.out.print("Number of records you want to enter into excel2 : ");
        n = sn.nextInt();
        for(int i=0;i<n;i++){
            addRecordToMap(map);
        }
        key = 0;

        addAllRecordsToSheet(sheet2,map);





        XSSFSheet sheet3 = wb.createSheet("All Students Data");
        map = new TreeMap<>();
        Integer key = 0;

        for(Row r : sheet1){
            Object[] obj = new Object[3];
            int i=0;
            for(Cell c : r){
                switch (c.getCellType()){
                    case Cell.CELL_TYPE_NUMERIC :
                        obj[i] = c.getNumericCellValue();
                        i++;
                        break;
                    case Cell.CELL_TYPE_STRING :
                        obj[i] = c.getStringCellValue();
                        i++;
                        break;
                }
            }
            System.out.println();
            map.put(key,obj);
            key++;
        }

        for(Row r : sheet2){
            Object[] obj = new Object[3];
            int i=0;
            Cell temp = r.getCell(0);
            if(temp.getCellType()==Cell.CELL_TYPE_STRING && temp.getStringCellValue()=="ID"){
                continue;
            }
            for(Cell c : r){
                switch (c.getCellType()){
                    case Cell.CELL_TYPE_NUMERIC:
                        obj[i] = c.getNumericCellValue();
                        i++;
                        break;
                    case Cell.CELL_TYPE_STRING :
                        obj[i] = c.getStringCellValue();
                        i++;
                        break;
                }
            }

            System.out.println();
            map.put(key,obj);
            key++;
        }



        addAllRecordsToSheet(sheet3,map);
        System.out.println();



        try{
            FileOutputStream fileOutputStream = new FileOutputStream("E:\\JavaFileHandling\\students.xlsx");
            wb.write(fileOutputStream);
            wb.close();
            System.out.println("successfully created the excel1 and excel2");
        }catch (Exception e){
            System.out.println("Some error occured");
        }

    }

    public static void addAllRecordsToSheet(XSSFSheet sheet,Map<Integer,Object[]> map){
        int rowNum = 0;
        for(Integer i : map.keySet()){
            Row r = sheet.createRow(rowNum++);
            int colNum = 0;
            for(Object obj : map.get(i)){
                Cell c = r.createCell(colNum++);
                if(obj instanceof  String){
                    c.setCellValue((String)obj);
                }else if (obj instanceof  Double){
                    c.setCellValue((Double) obj);
                }
            }
        }
    }


    public static void addRecordToMap(Map<Integer,Object[]> map){
        Scanner sn = new Scanner(System.in);
        System.out.print("Enter the id : ");
        double id = sn.nextDouble();
        sn.nextLine();
        System.out.print("Enter the Name : ");
        String name = sn.nextLine();
        System.out.print("Enter the marks : ");
        double marks = sn.nextDouble();
        map.put(++key,new Object[]{
                id,name,marks
        });

        System.out.println();

    }
}
