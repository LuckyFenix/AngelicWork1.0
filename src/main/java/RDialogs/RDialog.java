package RDialogs;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import javax.swing.*;
import javax.swing.table.JTableHeader;
import javax.swing.table.TableColumn;
import java.awt.*;
import java.io.FileInputStream;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collections;

public class RDialog extends JDialog
{
    public RDialog(Frame owner, String title, boolean modal)
    {
        super(owner, title, modal);
    }

    public int getResultSetSize(ResultSet rs) throws SQLException
    {
        rs.last();
        int max = rs.getRow();
        rs.beforeFirst();
        return max + 1;
    }

    public String smartTrim(String s)
    {
        String buff = s;
        s = "";
        for (int i = 0; i<buff.split(" ").length; i++)
        {
            if(i != 0)
                s = s + " ";
            s = s + buff.split(" ")[i];
        }
        buff = s;
        s = "";
        for (int i = 0; i<buff.split("'").length; i++)
        {
            s = s + buff.split("\\'")[i];
        }
        buff = s;
        s = "";
        for (int i = 0; i<buff.split("\"").length; i++)
        {
            s = s + buff.split("\\\"")[i];
        }
        return s;
    }

    public HSSFWorkbook readWorkbook(String filename)
    {
        try
        {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filename));
            return new HSSFWorkbook(fs);
        }
        catch (Exception e) {
            return null;
        }
    }

    public void setValueToSheet(HSSFSheet sheet, int rowNum, int cellNum, Double value)
    {
        HSSFRow row = sheet.getRow(rowNum);
        HSSFCell cell = row.getCell(cellNum);
        cell.setCellValue(value);
    }

    public void setValueToSheet(HSSFSheet sheet, int rowNum, int cellNum, String value)
    {
        HSSFRow row = sheet.getRow(rowNum);
        HSSFCell cell = row.getCell(cellNum);
        cell.setCellValue(value);
    }

    public void setFormulaToSheet(HSSFSheet sheet, int rowNum, int cellNum, String value)
    {
        HSSFRow row = sheet.getRow(rowNum);
        HSSFCell cell = row.getCell(cellNum);
        cell.setCellFormula(value);
    }

    public String getStringFromArray(ArrayList arrayList, char ch)
    {
        String s = "";
        for(int i = 0; i < arrayList.size(); i++)
        {
            if(i == 0)
                s += ch + arrayList.get(i).toString() + ch;
            else
                s += ", " + ch + arrayList.get(i).toString() + ch;
        }
        return s;
    }

    public String getStringFromArray(String[] arrayList, char ch)
    {
        String s = "";
        for(int i = 0; i < arrayList.length; i++)
        {
            if(i == 0)
                s += ch + arrayList[i] + ch;
            else
                s += ", " + ch + arrayList[i] + ch;
        }
        return s;
    }

    public ArrayList convertArray(Object[] s)
    {
        ArrayList arrayList = new ArrayList();
        Collections.addAll(arrayList, s);
        return arrayList;
    }

    public String[] convertArray(ArrayList<String> arrayList)
    {
        String[] s = new String[arrayList.size()];
        for (int i = 0; i < arrayList.size(); i++)
        {
            s[i] = arrayList.get(i);
        }
        return s;
    }

    public static void setColumnsWidth(JTable table)
    {
        table.setAutoResizeMode(JTable.AUTO_RESIZE_OFF);
        int prefWidth = 0;
        JTableHeader th = table.getTableHeader();
        for (int i = 0; i < table.getColumnCount(); i++)
        {
            TableColumn column = table.getColumnModel().getColumn(i);
            int prefWidthMax = 0;
            for (int j = 0; j <table.getRowCount(); j++)
            {
                String s = table.getModel().getValueAt(j, i).toString();
                prefWidth =
                        Math.round(
                                (float) th.getFontMetrics(
                                        th.getFont()).getStringBounds(s,
                                        th.getGraphics()
                                ).getWidth()
                        );
                if ( prefWidth > prefWidthMax ) prefWidthMax = prefWidth;
            }
            column.setPreferredWidth(prefWidthMax + 10);
        }
    }
}

