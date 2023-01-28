/*
 * Click nbfs://nbhost/SystemFileSystem/Templates/Licenses/license-default.txt to change this license
 * Click nbfs://nbhost/SystemFileSystem/Templates/Classes/Main.java to edit this template
 */
package javaapplication1;

import com.aspose.cells.Chart;
import com.aspose.cells.ChartType;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.opencsv.CSVParser;
import com.opencsv.CSVParserBuilder;
import com.opencsv.CSVReader;
import com.opencsv.CSVReaderBuilder;
import com.opencsv.exceptions.CsvException;
import java.io.BufferedReader;
import java.io.File;
import java.sql.*;
import java.io.FileNotFoundException;
import java.io.FileReader;
import java.io.IOException;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.ArrayList;
import java.util.List;
import java.util.Scanner;

/**
 *
 * @author mbuk
 */
public class JavaApplication1 {
    public static Path currentRelativePath = Paths.get("");
    public static String s = currentRelativePath.toAbsolutePath().toString();
    public static List<Games> games = new ArrayList<Games>();
    public static List<DataForDiagrams> data = new ArrayList<DataForDiagrams>();
    /**
     * @param args the command line arguments
     */
    public static void main(String[] args) throws FileNotFoundException, IOException, CsvException, ClassNotFoundException, SQLException, Exception {
       
       Class.forName("org.sqlite.JDBC");
//        System.out.println("Current absolute path is: " + s);
//        System.out.println("Hello World!");
        System.out.println("Введите пункт: 1. Парсинг; 2. Задание 2; 3. Задание 3; 4. Задание с диаграммой");
        switch(new Scanner(System.in).nextInt())
        {
            case 1:
                ReadCSV();
                break;
            case 2:
                System.out.println(GetTask2());
                break;
            case 3:
                System.out.println(GetTask3());
                break;
            case 4:
                GetTaskForDiagrams();
                break;
            default:
                break;
        }
    }
    
    public static void ReadCSV() throws FileNotFoundException, IOException, CsvException, SQLException{
        FileReader filereader = new FileReader(new File(s+"\\Games.csv"));
        CSVParser parser = new CSVParserBuilder().withSeparator(',').build();
        CSVReader csvReader = new CSVReaderBuilder(filereader)
                                  .withSkipLines(1)
                                  .withCSVParser(parser)
                                  .build();
        List<String[]> allData = csvReader.readAll();
       for (String[] row : allData) {
           Games game = new Games("N/A".equals(row[0]) ? -1 :Integer.parseInt(row[0]), row[1],row[2],"N/A".equals(row[3]) ? -1 : Integer.parseInt(row[3]), row[4],row[5],Double.parseDouble(row[6]),Double.parseDouble(row[7]),Double.parseDouble(row[8]),Double.parseDouble(row[9]),Double.parseDouble(row[10]));
           games.add((Games)game);
          
       }
       if(isEmptyDataInDb())
       {
           String sql = "INSERT INTO Games(Rank,Name,Platform,Year,Genre,Publisher,NA_Sales,EU_Sales,JP_Sales,Other_Sales,Global_Sales) VALUES(?,?,?,?,?,?,?,?,?,?,?)";
          Connection con = DriverManager.getConnection("jdbc:sqlite:"+s+"\\db.db");
          for(Games game : games){
              PreparedStatement stmt = con.prepareStatement(sql);
              stmt.setInt(1, game.Rank);
              stmt.setString(2, game.Name);
              stmt.setString(3, game.Platform);
              stmt.setInt(4, game.Year);
              stmt.setString(5, game.Genre);
              stmt.setString(6, game.Publisher);
              stmt.setDouble(7, game.NA_Sales);
              stmt.setDouble(8, game.EU_Sales);
              stmt.setDouble(9, game.JP_Sales);
              stmt.setDouble(10, game.Other_Sales);
              stmt.setDouble(11, game.Global_Sales);
              stmt.executeUpdate();
          }
            System.out.println("Данные успешно занесены в бд");
       }
       else{
           System.out.println("Данные уже есть в БД");
       }
    }
    
    public static boolean isEmptyDataInDb() throws SQLException{
        Connection con = DriverManager.getConnection("jdbc:sqlite:"+s+"\\db.db");
        Statement stmt = con.createStatement();
        int kol = 0;
        ResultSet rs = stmt.executeQuery("SELECT COUNT(*) AS kol FROM Games;");
        while ( rs.next() ){
            kol = rs.getInt("kol");
        }
        return kol == 0;
    }
    
    
    public static String GetTask2() throws SQLException{
        String sql = "SELECT s.Name as Name , MAX(s.Sum) as Max FROM (SELECT g.Name AS Name, SUM(g.EU_Sales) as Sum from Games g where g.Year = 2000 GROUP BY g.Name) s;";
        String name = "";
        Connection con1 = DriverManager.getConnection("jdbc:sqlite:"+s+"\\db.db");
        PreparedStatement  stmt = con1.prepareStatement(sql);
        ResultSet rs = stmt.executeQuery();
        
        while ( rs.next() ){
            name = rs.getString(1);
        }
        return name;
    }
    
    public static String GetTask3() throws SQLException{
        String sql = "SELECT s.Name as Name , MAX(s.Sum) as Max FROM (SELECT g.name AS Name, SUM(g.JP_Sales) as Sum from Games g where g.Year BETWEEN 2000 AND 2006 and g.Genre = \"Sports\" group by g.Name) s;";
        String name = "";
        Connection con1 = DriverManager.getConnection("jdbc:sqlite:"+s+"\\db.db");
        PreparedStatement  stmt = con1.prepareStatement(sql);
        ResultSet rs = stmt.executeQuery();
        while ( rs.next() ){
            name = rs.getString(1);
        }
        return name;
    }
    
    public static void GetTaskForDiagrams() throws Exception{
        
        if(!data.isEmpty()){
            data.clear();
        }
        
        String sql = "SELECT Platform, AVG(Global_Sales) as AVG_Sales FROM Games\n" +
            "GROUP BY Platform;";
        String platform = "";
        double avg_sales = 0.0;
        Connection con1 = DriverManager.getConnection("jdbc:sqlite:"+s+"\\db.db");
        PreparedStatement  stmt = con1.prepareStatement(sql);
        ResultSet rs = stmt.executeQuery();
        while ( rs.next() ){
            platform = rs.getString(1);
            avg_sales = rs.getDouble(2);
            data.add(new DataForDiagrams(platform, avg_sales));
        }
        
        Workbook workbook = new Workbook();
        Worksheet worksheet = workbook.getWorksheets().get(0);
        worksheet.getCells().get("A1").putValue("Platform");
        worksheet.getCells().get("B1").putValue("Avg_Sales");
        int k = 2;
        for(DataForDiagrams item : data){
            worksheet.getCells().get("A"+k).putValue(item.Platform);
            worksheet.getCells().get("B"+k).putValue(item.AVG_Sales);
            k++;
        }

        int chartIndex = worksheet.getCharts().add(ChartType.COLUMN, 5, 0, 15, 5);

        Chart chart = worksheet.getCharts().get(chartIndex);

        chart.setChartDataRange("A1:B"+k, true);
        workbook.save("Column-Chart.xlsx", SaveFormat.XLSX);
        System.out.println("Файл Column-Chart.xlsx сохранен в папке с проектом");
    }
}
