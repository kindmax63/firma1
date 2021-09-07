// Файл Firma1.class


package firma;  

import java.util.Iterator;

import java.util.Scanner;

import  java.io.*;

import  org.apache.poi.hssf.usermodel.HSSFSheet;

import  org.apache.poi.hssf.usermodel.HSSFWorkbook;

import  org.apache.poi.hssf.usermodel.HSSFRow;


//Программа по вводу информации о сотрудниках фирмы:

// Имя,фамилия, должность, возраст, оклад, стаж, парковочное место


class ParametrsCompanyTeam {

    String name;
    
    String secondname;
    
    String middlename;
    
    String position;
    
    int age;
    
    int oklad;
    
    int workexp;
    
    int parking;
    
                           }

public class Firma1 {

    public static void main(String[] args) {
    
        Scanner num = new Scanner(System.in);
        
        System.out.println("Введите количество сотрудников вашей фирмы:");
        
        int summarykol = num.nextInt();
        
        num.nextLine();
        
        ParametrsCompanyTeam[] sotr = new ParametrsCompanyTeam[summarykol];
        
        System.out.println("Вводите информацию о каждом сотруднике:");
        
        for (int i = 0; i < sotr.length; i++) {
        
        sotr[i] = new ParametrsCompanyTeam();
        

          System.out.println("Введите имя");
          
          sotr[i].name = num.nextLine();

          System.out.println("Введите фамилию");
          
          sotr[i].secondname= num.nextLine();

          System.out.println("Введите отчество");
          
          sotr[i].middlename= num.nextLine();

          System.out.println("Введите должность");
          
          sotr[i].position= num.nextLine();

          System.out.println("Введите возраст");
          
          sotr[i].age= num.nextInt();

          System.out.println("Введите оклад сотрудника");
          
          sotr[i].oklad= num.nextInt();

          System.out.println("Введите стаж работы в фирме (полных месяцев):");
          
          sotr[i].workexp= num.nextInt();

          System.out.println("Введите парковочное место, если его нет - введите ноль:");
          
          sotr[i].parking= num.nextInt();

          num.nextLine();
                                              }

          System.out.println( "\n Вывод информации по сотрудникам фирмы:");
          
          System.out.print("\n Имя \t Фамилия \t Отчество \t Должность \t Возраст \t Оклад");
          
          System.out.println("\t Стаж работы в фирме(месяцы) \t Номер парковочного места");

        for (ParametrsCompanyTeam s : sotr) {
            System.out.print(s.name +"\t"+s.secondname + "\t\t"+s.middlename
                    + "\t\t"+s.position + "\t\t"+s.age + "\t\t\t"+s.oklad + "\t\t\t\t"+s.workexp + "\t\t\t\t\t"+s.parking + "\n");
                                            }
        try {
            String filename = "C:/Users/Максим/Desktop/CompanyTeam.xlsx";
            HSSFWorkbook workbook = new HSSFWorkbook();
            HSSFSheet sheet = workbook.createSheet("FirstSheet");
            HSSFRow rowhead = sheet.createRow((short) 0);
            rowhead.createCell(0).setCellValue("Имя");
            rowhead.createCell(1).setCellValue("Фамилия");
            rowhead.createCell(2).setCellValue("Отчество");
            rowhead.createCell(3).setCellValue("Должность");
            rowhead.createCell(4).setCellValue("Возраст");
            rowhead.createCell(5).setCellValue("Оклад");
            rowhead.createCell(6).setCellValue("Стаж");
            rowhead.createCell(7).setCellValue("Парковочное место");

            for (ParametrsCompanyTeam s : sotr) {

                HSSFRow row = sheet.createRow((short) 1);

                row.createCell(0).setCellValue(s.name);
                row.createCell(1).setCellValue(s.secondname);
                row.createCell(2).setCellValue(s.middlename);
                row.createCell(3).setCellValue(s.position);
                row.createCell(4).setCellValue(s.age);
                row.createCell(5).setCellValue(s.oklad);
                row.createCell(6).setCellValue(s.workexp);
                row.createCell(7).setCellValue(s.parking);

            }


            FileOutputStream fileOut = new FileOutputStream(filename);
            workbook.write(fileOut);
            fileOut.close();
            workbook.close();
            System.out.println("Таблица сохранена в файл");
        }

        catch ( Exception ex ) {
            System.out.println(ex);
                                }


                                         }

                     }
