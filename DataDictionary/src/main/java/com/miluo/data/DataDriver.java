package com.miluo.data;

import com.miluo.data.dto.FieldDTO;
import com.miluo.data.dto.TableDTO;
import com.miluo.data.util.ExcelUtil;

import java.sql.*;
import java.util.ArrayList;
import java.util.List;

public class DataDriver {

    public static void main(String[] args) {
        String driver = "com.mysql.cj.jdbc.Driver";
        // 获取mysql连接地址
        String url = ""
        // 数据名称
        String username = "";
        // 数据库密码
        String password = "123456";
        // 获取一个数据的连接
        Connection conn = null;

        try{
            Class.forName(driver);
            //getConnection()方法，连接MySQL数据库！
            conn= DriverManager.getConnection(url,username,password);
            ResultSet tableRet = conn.getMetaData().getTables("feishou_test2", "%","%",new String[]{"TABLE"});
            List tables = new ArrayList<TableDTO>();

            while(tableRet.next()) {
//                if (tableRet.getString("TABLE_NAME").indexOf("seq") == 0 || tableRet.getString("TABLE_NAME").indexOf("da") == 0) {
//                    continue;
//                }

                if (tableRet.getString("TABLE_NAME").indexOf("da") != 0) {
                    continue;
                }
                System.out.println(tableRet.getString("TABLE_NAME"));
                TableDTO table = new TableDTO();
                table.setTableName(tableRet.getString("TABLE_NAME"));
                table.setTableRemark(tableRet.getString("REMARKS"));
                ResultSet colRet = conn.getMetaData().getColumns("feishou_test2", "%", tableRet.getString("TABLE_NAME"), "%");
                ResultSet primaryKeyResultSet = conn.getMetaData().getPrimaryKeys("feishou_test2", null, tableRet.getString("TABLE_NAME"));
                ResultSet foreignKeyResultSet = conn.getMetaData().getImportedKeys("feishou_test2", null, tableRet.getString("TABLE_NAME"));

                List primaryKey = new ArrayList<String>();
                while (primaryKeyResultSet.next()) {
                    primaryKey.add(primaryKeyResultSet.getString("COLUMN_NAME"));
                }

                List foreignKey = new ArrayList<String>();
                while (foreignKeyResultSet.next()) {
                    foreignKey.add(foreignKeyResultSet.getString("FKCOLUMN_NAME"));
                }

                List fields = new ArrayList<FieldDTO>();
                while (colRet.next()) {
                    FieldDTO field = new FieldDTO();
                    String columnName = colRet.getString("COLUMN_NAME");
                    field.setColumnName(colRet.getString("REMARKS"));
                    field.setColumnCode(columnName);
                    field.setDataType(colRet.getString("TYPE_NAME"));
                    field.setSize(colRet.getInt("COLUMN_SIZE"));
                    if (primaryKey.indexOf(columnName) >= 0) {
                        field.setPK("Y");
                    }
                    if (foreignKey.indexOf(columnName) >= 0) {
                        field.setFK("Y");
                    }
                    fields.add(field);
                }
                table.setFields(fields);
                System.out.println(table);
                tables.add(table);
            }
            conn.close();
            System.out.println(tables.size());
            ExcelUtil.getDictionary(tables);
        }
        catch(ClassNotFoundException e){
            //数据库驱动类异常处理
            System.out.println("数据库驱动加载失败！");
            e.printStackTrace();
        }
        catch(SQLException e1){
            //数据库连接失败异常处理
            e1.printStackTrace();
        }
        catch(Exception e2){
            e2.printStackTrace();
        }
        finally{
            System.out.println("-------------------------------");
            System.out.println("数据库数据获取成功！");
        }


    }
}
