package exportexcel;

import java.io.BufferedReader;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.InputStream;
import java.sql.Connection;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.Statement;
import java.util.Properties;

import org.apache.commons.codec.binary.Hex;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
//import java.awt.Color;

public class ExportToExcel {
    public static void main(String[] args) {
        if (args.length != 1) {
            System.out.println("Usage: java ExportToExcel <config_file_path>");
            return;
        }

        String configFile = args[0];

        Properties prop = new Properties();
        try (InputStream input = new FileInputStream(configFile)) {
            prop.load(input);

            String url = prop.getProperty("db.url");
            String username = prop.getProperty("db.username");
            String password = prop.getProperty("db.password");
            String sqlQueryFilePath = prop.getProperty("db.sqlQuery");
            String outputPath = prop.getProperty("output.path");
            String headerColorCode = prop.getProperty("excel.HeadColor");

            byte[] rgbB = Hex.decodeHex(headerColorCode);

            StringBuilder sqlQueryBuilder = new StringBuilder();
            try (BufferedReader reader = new BufferedReader(new FileReader(sqlQueryFilePath))) {
                String line;
                while ((line = reader.readLine()) != null) {
                    sqlQueryBuilder.append(line).append("\n");
                }
            }

            String sqlQuery = sqlQueryBuilder.toString().trim();
            try (
                Connection connection = DriverManager.getConnection(url, username, password);
                Statement statement = connection.createStatement();
                ResultSet resultSet = statement.executeQuery(sqlQuery);
                Workbook workbook = new XSSFWorkbook();
            ) {
                Sheet sheet = workbook.createSheet("Data");

                ResultSetMetaData metaData = resultSet.getMetaData();
                int columnCount = metaData.getColumnCount();
                XSSFColor color = new XSSFColor(rgbB, null);

                XSSFCellStyle cellStyle = (XSSFCellStyle) workbook.createCellStyle();
                cellStyle.setFillForegroundColor(color);
                cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                cellStyle.setBorderTop(BorderStyle.THIN);
                cellStyle.setBorderBottom(BorderStyle.THIN);
                cellStyle.setBorderLeft(BorderStyle.THIN);
                cellStyle.setBorderRight(BorderStyle.THIN);

                // Create a Font for the header text
                Font headerFont = workbook.createFont();
                headerFont.setBold(true);
                cellStyle.setFont(headerFont);

                Row headerRow = sheet.createRow(0);
                for (int i = 1; i <= columnCount; i++) {
                    Cell cell = headerRow.createCell(i - 1);
                    cell.setCellValue(metaData.getColumnName(i));
                    cell.setCellStyle(cellStyle);
                }
                CellStyle cellStyle1 = workbook.createCellStyle();

                // Apply a thin border to all sides of the cell
                cellStyle1.setBorderTop(BorderStyle.THIN);
                cellStyle1.setBorderBottom(BorderStyle.THIN);
                cellStyle1.setBorderLeft(BorderStyle.THIN);
                cellStyle1.setBorderRight(BorderStyle.THIN);

                int rowNum = 1;
                while (resultSet.next()) {
                    Row dataRow = sheet.createRow(rowNum++);
                    for (int i = 1; i <= columnCount; i++) {
                        Cell cell = dataRow.createCell(i - 1);
                        cell.setCellValue(resultSet.getString(i));
                        cell.setCellStyle(cellStyle1);

                    }
                }

                try (FileOutputStream outputStream = new FileOutputStream(outputPath)) {
                    workbook.write(outputStream);
                    System.out.println("Data exported to " + outputPath);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
