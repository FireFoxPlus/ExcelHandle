package com.picc.cdyf.utils;

import java.io.FileInputStream;
import java.io.InputStream;
import java.text.DecimalFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.hibernate.Hibernate;
import org.hibernate.SQLQuery;
import org.hibernate.Session;
import org.hibernate.SessionFactory;
import org.hibernate.cfg.Configuration;

public class MyExcelUtils {
	public int dataSaveToDB(List<List<String>> rs){
	    SessionFactory sessionFactory = new Configuration().configure().buildSessionFactory();
	    Session session = sessionFactory.getCurrentSession();
		session.beginTransaction();
		for(int i = 0; i < rs.size(); i++){
	       List<String> row = rs.get(i);
		   StringBuffer sql = new StringBuffer();
           sql.append("insert into gk values(?, ?, ?, ? , ?, ?, ?, ?, ?, ?)");
          // sql.append("insert into gk values('s', 's', 's', 'd', '4', 't', 'i', 'o', 'p', 'e')");
		   SQLQuery sqlQuery = session.createSQLQuery(sql.toString());
		   for(int j = 0; j < row.size(); j++){
			   sqlQuery.setParameter(j, row.get(j));
		   }
	       sqlQuery.executeUpdate();
	    } 
		session.getTransaction().commit();
		
		
		
		return 0;
	}
	

	@SuppressWarnings("static-access")
	public List<List<String> > readXls(String path) throws Exception{
		InputStream is = new FileInputStream(path);
		HSSFWorkbook hssfWorkbook = new HSSFWorkbook(is);
		List<List<String> > rs = new ArrayList<List<String>>();
		//循环处理每一页
		for(int numSheet = 0; numSheet < hssfWorkbook.getNumberOfSheets(); numSheet++){
			HSSFSheet hssfSheet = hssfWorkbook.getSheetAt(numSheet);
			if(hssfSheet == null)
				continue;
			for(int rowNum = 1; rowNum <= hssfSheet.getLastRowNum(); rowNum++){
				HSSFRow hssfRow = hssfSheet.getRow(rowNum);
				if(hssfRow == null)
					break;
				int minColIx = hssfRow.getFirstCellNum();
				int maxColIx = hssfRow.getLastCellNum();
				List<String> rowList = new ArrayList<String>();
				//遍历该行，获取cell
				for(int colIx = minColIx; colIx < maxColIx; colIx++){
					HSSFCell cell = hssfRow.getCell(colIx);
					String value = "";
					if (cell != null) {
		            // 注意：一定要设成这个，否则可能会出现乱码
		                //cell.setEncoding(HSSFCell.e);
		                switch (cell.getCellType()) {
		                     case HSSFCell.CELL_TYPE_STRING:
		                         value = cell.getStringCellValue();
		                         break;
		                     case HSSFCell.CELL_TYPE_NUMERIC:
		                         if (HSSFDateUtil.isCellDateFormatted(cell)) {
		                            Date date = cell.getDateCellValue();
		                            if (date != null) {
		                                value = new SimpleDateFormat("yyyy-MM-dd").format(date);
		                            } else {
		                                value = "";
		                            }
		                         } else {
		                            value = new DecimalFormat("0").format(cell.getNumericCellValue());
		                         }
		                         break;
		                     case HSSFCell.CELL_TYPE_FORMULA:
		                         // 导入时如果为公式生成的数据则无值
		                         if (!cell.getStringCellValue().equals("")) {
		                            value = cell.getStringCellValue();
		                         } else {
		                            value = cell.getNumericCellValue() + "";
		                         }
		                         break;
		                     case HSSFCell.CELL_TYPE_BLANK:
		                         break;
		                     case HSSFCell.CELL_TYPE_ERROR:
		                         value = "";
		                         break;
		                     case HSSFCell.CELL_TYPE_BOOLEAN:
		                         value = (cell.getBooleanCellValue() == true ? "Y"  : "N");
		                         break;
		                     default:
		                         value = "";
		                     }
		                  }
		           	if(cell == null)
						continue;
					rowList.add(value);
				   }
				rs.add(rowList);
				}	
		}
		return rs;
	}

}
