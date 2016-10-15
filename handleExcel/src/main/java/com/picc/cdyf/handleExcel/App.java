package com.picc.cdyf.handleExcel;

import java.util.List;

import com.picc.cdyf.utils.MyExcelUtils;

/**
 * Hello world!
 *
 */
public class App {
    public static void main( String[] args ){
       MyExcelUtils myExcelUtils = new MyExcelUtils();
       List<List<String>> rs = null;
       try {
    	   rs = myExcelUtils.readXls("e:\\data.xls");
       } catch (Exception e) {
    	   e.printStackTrace();
	   }
       myExcelUtils.dataSaveToDB(rs);
//       for(int i = 0; i < rs.size(); i++){
//           List<String> row = rs.get(i);
//           for(int j = 0; j < row.size(); j++){
//        	   System.out.print(row.get(j) + "|");
//           }
//    	   System.out.println();
//       }   
    }
}
