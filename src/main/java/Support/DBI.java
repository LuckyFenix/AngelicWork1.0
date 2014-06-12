package Support;

import java.sql.*;
import java.util.ArrayList;
import java.util.Properties;

public class DBI
{
    private Connection bd;
    private Statement st;
    private String dbName;
    private String url;

    public DBI(String dataBaseName)
    {
        /*Properties properties = new Properties();

        InputStream is = this.getClass().getResourceAsStream( "db_conection.properties" );

        File propertyFile = new File(getClass().getResource("db_conection.properties").getFile());

        try
        {
            properties.load(is);
        } catch (IOException e) {
            e.printStackTrace();
        }*/

        this.dbName = dataBaseName;

        Properties connInfo = new Properties();

        connInfo.put("user", "root"/*properties.getProperty("DataBaseUser")*/);
        connInfo.put("password", "root"/*properties.getProperty("DataBasePass")*/);
        connInfo.put("charSet", "UTF8");
        try
        {

            Class.forName("com.mysql.jdbc.Driver"/*properties.getProperty("DataBaseDriverName").trim()*/);
            url = "jdbc:mysql://" + "localhost"/*properties.getProperty("DataBaseHost")*/ + "/" + dbName + "?useUnicode=true&characterEncoding=utf8";
            bd = DriverManager.getConnection(url, connInfo);
            st = bd.createStatement();
        } catch (ClassNotFoundException ex)
        {
            ex.printStackTrace();
        } catch (SQLException e)
        {
            try
            {
                url = "jdbc:mysql://" + "localhost"/*properties.getProperty("DataBaseHost")*/ + "/information_schema";
                bd = DriverManager.getConnection(url, connInfo);
                st = bd.createStatement();
                st.execute("CREATE DATABASE " + dbName + " CHARACTER SET utf8 COLLATE utf8_general_ci;");
                url = "jdbc:mysql://" + "localhost"/*properties.getProperty("DataBaseHost")*/ + "/" + dbName + "?useUnicode=true&characterEncoding=utf8";
                bd = DriverManager.getConnection(url, connInfo);
                st = bd.createStatement();
            } catch (SQLException e1)
            {
                e1.printStackTrace();
            }
        }
    }

    public boolean createTable(String tableName, String[][] colunm)
    {
        try
        {
            StringBuilder request = new StringBuilder(100);
            request.append("create table if not exists '")
                    .append(tableName)
                    .append("' (");
            for (String[] aColunm : colunm)
            {
                request.append("'")
                        .append(aColunm[0])
                        .append("' ")
                        .append(aColunm[1])
                        .append(", ");
            }
            request.delete(request.length() - 2, request.length());
            request.append(");");
            st.execute(request.toString());
            return true;
        }
        catch (SQLException ex)
        {
            System.err.println(ex);
            return false;
        }
    }

    public String[][] getStringArray(String tableName)
    {
        ArrayList<ArrayList<String>> list = new ArrayList<>();
        list.add(new ArrayList<>());
        String[][] result = null;
        try
        {
            ResultSet rs = st.executeQuery("select * from " + tableName);
            for(int i = 0; rs.next(); i++)
            {
                list.add(new ArrayList<>());
                for(int j = 1; j<100000; j++)
                {
                    try
                    {
                        list.get(i).add(rs.getString(j));
                    } catch (SQLException ex)
                    {
                        break;
                    }
                }
            }
            result = new String[list.size() - 1][list.get(1).size()];
            for(int i = 0; i<result.length; i++)
            {
                for(int j = 0; j<result[i].length; j++)
                {
                    result[i][j] = list.get(i).get(j);
                }
            }
        }
        catch (SQLException ex)
        {
            System.err.println(ex);
            return null;
        }
        return result;
    }

    public int[][] getIntegerArray(String tableName)
    {
        ArrayList<ArrayList<String>> list = new ArrayList<>();
        list.add(new ArrayList<>());
        int[][] result = null;
        try
        {
            ResultSet rs = st.executeQuery("select * from " + tableName);
            for(int i = 0; rs.next(); i++)
            {
                list.add(new ArrayList<>());
                for(int j = 1; j<100000; j++)
                {
                    try
                    {
                        list.get(i).add(rs.getString(j));
                    } catch (SQLException ex)
                    {
                        break;
                    }
                }
            }
            result = new int[list.size() - 1][list.get(1).size()];
            for(int i = 0; i<result.length; i++)
            {
                for(int j = 0; j<result[i].length; j++)
                {
                    result[i][j] = Integer.parseInt(list.get(i).get(j));
                }
            }
            rs.close();
        }
        catch (SQLException ex)
        {
            System.err.println(ex);
            return result;
        }
        return result;
    }

    public boolean insertIntegerArray(String tableName, int[][] intArray)
    {
        try {
            StringBuilder sBuffer = new StringBuilder();
            ResultSet rs = st.executeQuery("select * from " + tableName);
            ResultSet columns = st.executeQuery("PRAGMA table_info('" + tableName + "')");
            ArrayList<String> columnsName = new ArrayList<>();
            while(columns.next())
            {
                columnsName.add(columns.getString("name"));
            }
            for (int[] anIntArray : intArray)
            {
                sBuffer.append("insert into '")
                        .append(tableName)
                        .append("' (");
                for (int j = 0; j < intArray[0].length; j++)
                {
                    sBuffer.append("'")
                            .append(columnsName.get(j))
                            .append("', ");
                }
                sBuffer.delete(sBuffer.length() - 2, sBuffer.length());
                sBuffer.append(") values (");

                for (int j = 0; j < intArray[0].length; j++)
                {
                    sBuffer.append("")
                            .append(anIntArray[j])
                            .append(", ");
                }
                sBuffer.delete(sBuffer.length() - 2, sBuffer.length());
                sBuffer.append(")");
                st.execute(sBuffer.toString());
                sBuffer.delete(0, sBuffer.length());
            }
        } catch (SQLException ex)
        {
            System.err.println(ex);
            return false;
        }
        return true;
    }

    public boolean delete(String tableName)
    {
        try
        {
            ResultSet columns = st.executeQuery("PRAGMA table_info('" + tableName + "')");
            ArrayList<String> columnsName = new ArrayList<>();
            while(columns.next())
            {
                columnsName.add(columns.getString("name"));
            }
            st.execute("delete from " + tableName);
        } catch (SQLException ex) {
            System.err.println(ex);
            return false;
        }
        return true;
    }

    public void close(int k)
    {
        try
        {
            k--;
            bd.close();
            st.close();
        } catch (SQLException ex) {
            ex.printStackTrace();
        }
    }

    public Connection getBd() {
        return bd;
    }

    public Statement getSt() {
        return st;
    }

    public String getDbName() {
        return dbName;
    }
}
