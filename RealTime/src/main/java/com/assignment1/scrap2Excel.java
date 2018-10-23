package com.assignment1;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class scrap2Excel {

    public static List<Data> data=new ArrayList();


    public static void scrapData(){
        try{
            System.out.println("Accessing...");

            Document source= Jsoup.connect("https://ms.wikipedia.org/wiki/Malaysia").get();
            Element table=source.select("table.wikitable").get(1);
            Elements rows=table.select("tr");

            for (Element row:rows){
                Elements data1=row.select("th");
                Elements data2=row.select("td");
                String coloumn1=data1.text();
                String coloumn2=data2.text();

                data.add(new Data(coloumn1,coloumn2));
            }

            for (int i=0;i<=23;i++){
                String output1=data.get(i).getData1();
                System.out.printf("%40s :",output1);
                System.out.println(data.get(i).getData2());
            }


        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void write2Excel(){

        XSSFWorkbook workbook=new XSSFWorkbook();
        XSSFSheet sheet=workbook.createSheet("Trivia");
        System.out.println("Writing to excel file...");

        try {
            for (int i=0;i<data.size();i++){
                XSSFRow row=sheet.createRow(i);

                XSSFCell cell=row.createCell(0);
                cell.setCellValue(data.get(i).getData1());
                XSSFCell cell2=row.createCell(1);
                cell2.setCellValue(data.get(i).getData2());
            }
            FileOutputStream fileOutputStream = new FileOutputStream("C:\\Users\\Koon Fung Yee\\Desktop\\Malaysia.xlsx");
            workbook.write(fileOutputStream);
            workbook.close();
            System.out.println("done");

        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    public static void main(String[] args) {
        scrapData();
        write2Excel();
    }

}
