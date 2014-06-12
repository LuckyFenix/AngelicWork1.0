package RDialogs;

import Main.Main;
import Support.DBI;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellStyle;

import javax.swing.*;
import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Iterator;

public class ExportDialog
{
    private String category;
    private String group;

    public ExportDialog(Main main) throws SQLException
    {
        DBI dbi = new DBI("databassesabc");
        ResultSet rs = dbi.getSt().executeQuery("SELECT `Категорія` FROM `товари` GROUP BY BINARY `Категорія`;");
        ArrayList<String> categoryArray = new ArrayList<>();
        while (rs.next())
        {
            String s = rs.getString(1);
            if (s.charAt(0)<='я' && s.charAt(0)>='а')
            {
                DBI dbiU = new DBI("databassesabc");
                dbiU.getSt().execute("UPDATE `товари` " +
                        "SET `Категорія`='" + (Character.toUpperCase(s.charAt(0)) + s.substring(1, s.length())) + "' " +
                        "WHERE `Категорія`='" + s +"';");
                dbiU.close(main.k);
            }
            categoryArray.add(s);
        }
        rs.close();
        dbi.close(main.k);

        JFileChooser fileChooser = new JFileChooser(new File("."));
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int result = fileChooser.showSaveDialog(main);

        for (String aCategoryArray : categoryArray)
        {
            if (result == JFileChooser.APPROVE_OPTION)
            {
                String filename = fileChooser.getSelectedFile().getPath();

                ArrayList<String> groupsArray = new ArrayList<>();
                groupsArray.clear();
                dbi = new DBI("databassesabc");
                rs = dbi.getSt().executeQuery("SELECT `Група` FROM `товари` WHERE `Категорія`='" + aCategoryArray + "' GROUP BY `Група` ORDER BY `Група` ASC;");
                while (rs.next())
                {
                    groupsArray.add(rs.getString(1));
                }
                rs.close();
                dbi.close(main.k);

                HSSFWorkbook workbook = readWorkbook("Товари.xls");
                HSSFSheet sheet = workbook.getSheetAt(0);
                Iterator rows = sheet.rowIterator();
                HSSFRow row;
                HSSFCell cell;

                rows.next();
                rows.next();

                int k = 1;
                for (String aGroupsArray : groupsArray)
                {
                    dbi = new DBI("databassesabc");
                    String[] column = new String[]{"Категорія", "Група", "Назва товару", "Залишок на початок. Склад", "Залишок на початок. Магазин", "Прихід", "Ціна", "Дата останнього приходу"};
                    if (aGroupsArray.equals("Годинник, фонарик, магнітики") || aGroupsArray.equals("Машинки"))
                        rs = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `товари` WHERE `Група`='" + aGroupsArray + "' ORDER BY `Ціна` ASC;");
                    else
                        rs = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `товари` WHERE `Група`='" + aGroupsArray + "' ORDER BY `Назва товару` ASC;");


                    while (rs.next())
                    {
                        String newCategory = rs.getString(1);
                        String newGroup = rs.getString(2);


                        if (!newCategory.equals(category))
                        {
                            rows.hasNext();
                            row = (HSSFRow) rows.next();

                            CellStyle style = workbook.createCellStyle();
                            style.setFillForegroundColor(HSSFColor.RED.index);
                            style.setFillPattern(CellStyle.SOLID_FOREGROUND);

                            cell = row.getCell(1);
                            cell.setCellValue(newCategory);
                            for (int j = 0; j < 7; j++)
                            {
                                row.getCell(j).setCellStyle(style);
                            }

                            category = newCategory;
                        }
                        if (!newGroup.equals(group))
                        {
                            rows.hasNext();
                            row = (HSSFRow) rows.next();

                            CellStyle style = workbook.createCellStyle();
                            style.setFillForegroundColor(HSSFColor.LIGHT_BLUE.index);
                            style.setFillPattern(CellStyle.SOLID_FOREGROUND);

                            cell = row.getCell(1);
                            cell.setCellValue(newGroup);
                            for (int j = 0; j < 7; j++)
                            {
                                row.getCell(j).setCellStyle(style);
                            }

                            group = newGroup;
                        }
                        rows.hasNext();
                        row = (HSSFRow) rows.next();

                        row.getCell(0).setCellValue(k);
                        row.getCell(1).setCellValue(rs.getString(3));
                        row.getCell(2).setCellValue(rs.getString(4));
                        row.getCell(3).setCellValue(rs.getString(5));
                        row.getCell(4).setCellValue(rs.getString(6));
                        row.getCell(5).setCellValue(rs.getString(7));

                        try
                        {
                            row.getCell(6).setCellValue(main.getDateFormat().format(rs.getDate(8)));
                        } catch (SQLException ignored)
                        {
                        }
                        k++;
                    }
                    rs.close();
                    dbi.close(main.k);
                }

                try
                {
                    FileOutputStream fileOut = new FileOutputStream(filename + "\\" + aCategoryArray + ".xls");
                    workbook.write(fileOut);
                    fileOut.close();
                } catch (IOException e1)
                {
                    e1.printStackTrace();
                }

                //try
                //{
                //Runtime.getRuntime().exec("cmd /c start excel \" " + filename + "\\Товар.xls\"");
                //} catch (IOException e1) {}
            }
        }
    }

    public HSSFWorkbook readWorkbook(String filename)
    {
        try
        {
            POIFSFileSystem fs = new POIFSFileSystem(new FileInputStream(filename));
            return new HSSFWorkbook(fs);
        } catch (Exception e) {return null;}
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

    public void setValueToSheet(HSSFSheet sheet, int rowNum, int cellNum, String value)
    {
        HSSFRow row = sheet.getRow(rowNum);
        HSSFCell cell = row.getCell(cellNum);
        cell.setCellValue(value);
    }
}
