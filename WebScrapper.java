package com.crm.qa.base;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;
import org.jsoup.nodes.Element;
import org.jsoup.select.Elements;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

public class Announcements {
    public static void main(String args[]) {
        //pass the url as the input argument.
        try {
            File file = new File("/Users/archanagupta/Desktop/Announcements-Upstox.html");
            Document document = Jsoup.parse(file,"UTF-8");
            String title = document.title();
            System.out.println(title);
            Element body = document.body();
            Element pageContainer = body.getElementsByClass("page-container").first();
            Elements btns = document.getElementsByClass("a-article-wrapper");
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFSheet spreadsheet = workbook.createSheet(" Announcements Data ");
            XSSFRow xsRow;
            int rowid = 0;
            for(Element btn : btns){
                Elements rows = btn.getElementsByClass("a-article-row");
                for(Element row:rows) {
                    xsRow = spreadsheet.createRow(rowid++);
                    int cellid = 0;
                    String currentDate = row.getElementsByClass("a-current-date").first().text();
                    String currentTime = row.getElementsByClass("a-current-time").first().text();
                    String articleTitle = row.getElementsByClass("a-article-title").first().text();
                    String tags = row.getElementsByClass("a-article-tags").first().text();
                    Element aContent = row.getElementsByClass("a-article-content").first();
                    String content = aContent.getElementsByClass("content").first().html();
                    List<String> data = new ArrayList<String>();
                    data.add(currentDate);
                    data.add(currentTime);
                    data.add(articleTitle);
                    data.add(tags);
                    data.add(content);
                    for (String obj : data) {
                        Cell cell = xsRow.createCell(cellid++);
                        cell.setCellValue(obj);
                    }
                }
            }
            FileOutputStream out = new FileOutputStream(new File("/Users/archanagupta/Desktop/announcements-data1.xlsx"));
            workbook.write(out);
            out.close();
        } catch (Exception e) {
            e.getMessage();
        }
    }

    private static void createExcel() {
        XSSFWorkbook workbook = new XSSFWorkbook();
        XSSFSheet spreadsheet = workbook.createSheet(" Announcements Data ");
        XSSFRow row;
    }
}
