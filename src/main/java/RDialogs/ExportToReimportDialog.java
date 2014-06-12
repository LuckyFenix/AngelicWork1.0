package RDialogs;

import Main.Main;
import Support.DBI;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import javax.swing.*;
import java.io.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.Iterator;


public class ExportToReimportDialog
{
    private String filename;
    private HSSFWorkbook workbook;

    public ExportToReimportDialog(Main main) throws SQLException
    {
        JFileChooser fileChooser = new JFileChooser(new File("."));
        fileChooser.setFileSelectionMode(JFileChooser.DIRECTORIES_ONLY);

        int result = fileChooser.showSaveDialog(main);


        if (result == JFileChooser.APPROVE_OPTION)
        {
            filename = fileChooser.getSelectedFile().getPath();

            workbook = readWorkbook("Товари2.xls");
            HSSFSheet sheet = workbook.getSheetAt(0);
            Iterator rows = sheet.rowIterator();
            HSSFRow row;

            rows.next();
            rows.next();

            int k = 1;

            DBI dbi = new DBI("databassesabc");
            String[] column = new String[]{"Категорія", "Група", "Назва товару", "Одиниці вимірювання", "Товару в загальному", "Ціна закупочна", "Ціна"};
            ResultSet rs = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `товари` ORDER BY `Категорія`, `Група`, `Назва товару` ASC;");
            while (rs.next())
            {
                rows.hasNext();
                row = (HSSFRow) rows.next();

                row.getCell(0).setCellValue(k);
                for (int i = 1; i < 5; i++)
                {
                    row.getCell(i).setCellValue(rs.getString(i));
                }
                for (int i = 5; i < 8; i++)
                {
                    row.getCell(i).setCellValue(rs.getDouble(i));
                }
                k++;
            }
            rs.close();
            dbi.close(main.k);
        }

        try
        {
            FileOutputStream fileOut = new FileOutputStream(filename + "\\Файл імпорта.xls");
            workbook.write(fileOut);
            fileOut.close();
        } catch (FileNotFoundException e1) {e1.printStackTrace();} catch (IOException e1) {e1.printStackTrace();}
    }

    private HSSFWorkbook readWorkbook(String filename)
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
}
