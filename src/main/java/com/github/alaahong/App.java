package com.github.alaahong;

import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.DriverManager;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.sql.Statement;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import com.mysql.jdbc.Connection;

/**
 * Hello world!
 * 
 */
public class App {
	public static void main(String[] args) throws IOException {
		List lineList = new ArrayList();
		try {
			Class.forName("com.mysql.jdbc.Driver");
			String url = "jdbc:mysql://192.168.30.180:3306/perflog?useUnicode=true&characterEncoding=utf-8";
			String user = "root";
			String password = "mysql";
			Connection con = (Connection) DriverManager.getConnection(url,
					user, password);
			Statement statement = con.createStatement();
			String querySql = "";
			if (args.length == 0) {
				System.out.println("please input the date!!!");
				System.exit(-1);
			} else {
				querySql = "select r.* from record as r where edate = '"
						+ args[0].trim()
						+ "' and 10> (select count(*) from record where edate = '"
						+ args[0].trim()
						+ "' and module = r.module AND stime > r.stime LIMIT 0,10) ORDER BY r.module,r.stime desc;";
			}
			ResultSet rs = statement.executeQuery(querySql);
			while (rs.next()) {
				List<String> cellList = new ArrayList<String>();
				cellList.add(rs.getString("eid"));
				cellList.add(rs.getString("module"));
				cellList.add(rs.getString("stime"));
				cellList.add(rs.getString("edate"));
				cellList.add(rs.getString("fname"));
				cellList.add(rs.getString("event"));
				cellList.add(rs.getString("reqs"));
				lineList.add(cellList);
			}
			rs.close();
			statement.close();
			con.close();

		} catch (ClassNotFoundException e) {
			e.printStackTrace();
		} catch (SQLException e) {
			e.printStackTrace();
		}

		// 导出xls格式的excel
		// Workbook wb = new HSSFWorkbook();
		// 导出xlsx格式的excel
		Workbook wb = new XSSFWorkbook();
		DataFormat format = wb.createDataFormat();
		CellStyle style;
		// 创建一个SHEET
		Sheet sheet1 = wb.createSheet("Top10AsModule");
		// // 设置表头要显示的内容
		// String[] title = { "eid", "module", "stime", "edate", "fname",
		// "event", "reqs" };
		// int i = 0;
		// // 创建一行
		// Row row = sheet1.createRow((short) 0);
		// // 填充标题.将标题放入第一行各列
		// for (String s : title) {
		// Cell cell = row.createCell(i);
		// cell.setCellValue(s);
		// i++;
		// }
		for (int s = 0; s < lineList.size(); s++) {
			List<String> cellList = (List<String>) lineList.get(s);
			Row row1 = sheet1.createRow((short) s);
			row1.createCell(0).setCellValue(cellList.get(0).toString());
			row1.createCell(1).setCellValue(cellList.get(1).toString());
			row1.createCell(2).setCellValue(cellList.get(2).toString());
			row1.createCell(3).setCellValue(cellList.get(3).toString());
			row1.createCell(4).setCellValue(cellList.get(4).toString());
			row1.createCell(5).setCellValue(cellList.get(5).toString());
			row1.createCell(6).setCellValue(cellList.get(6).toString());
		}
		FileOutputStream fileOut = new FileOutputStream(args[0] + ".xlsx");
		wb.write(fileOut);
		fileOut.close();
		System.out
		.println(args[0]+".xlsx");
	}
}
