package Waybills;

import Main.Main;
import RComponents.*;
import Support.DBI;
import Support.GBC;
import com.michaelbaranov.microba.calendar.DatePicker;
import org.apache.poi.hssf.usermodel.*;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;

public class Revision extends RPanel
{
    private JPanel panel = new JPanel(new GridBagLayout());
    private DatePicker datePicker;
    private final Main main;
    private DBI dbi;
    private ResultSet resultSet;
    private String[] column;
    private DefaultTableModel tableModel;
    private JTable table;
    private ArrayList<JButton> btnArray = new ArrayList<>();
    private ArrayList<String> category = new ArrayList<>();

    public Revision(Main main)
    {
        this.main = main;

        setLayout(new GridBagLayout());

        initPanel();
    }

    private void initPanel()
    {
        this.removeAll();

        datePicker = new DatePicker();
        datePicker.addActionListener(e ->
        {
            initValue();
            initTable();
        });

        initColumn();
        table = new JTable(tableModel);
        initTable();
        initValue();

        String fileName;
        JButton importBtn = new JButton("Імпортувати з Excel");
        importBtn.addActionListener(new ImportBtnListener());
        try
        {
            dbi = new DBI("category");
            resultSet = dbi.getSt().executeQuery("SHOW TABLES FROM `category`;");
            for (int i = 0; resultSet.next(); i++)
            {
                fileName = resultSet.getString(1);
                fileName = Character.toUpperCase(fileName.charAt(0)) + fileName.substring(1, fileName.length());
                btnArray.add(new JButton(fileName));
                category.add(fileName);
                final String finalFileName = fileName;
                btnArray.get(i).addActionListener(e ->
                {
                    try
                    {
                        Runtime.getRuntime().exec("cmd /c start excel \"" + finalFileName + "\"");
                    } catch (IOException e1) {e1.printStackTrace();}
                });
            }
        } catch (SQLException e) {e.printStackTrace();}

        JScrollPane scrollPane = new JScrollPane(table);
        scrollPane.setPreferredSize(new Dimension(1200, 400));
        panel.add(datePicker, new GBC(0, 0).setAnchor(GridBagConstraints.NORTHWEST));
        panel.add(scrollPane, new GBC(1, 0, 1, btnArray.size() + 3).setAnchor(GridBagConstraints.NORTHWEST).setInsets(0, 5, 0, 0));
        panel.add(importBtn, new GBC(0, 1).setAnchor(GridBagConstraints.CENTER));
        for (int i = 0; i < btnArray.size(); i++)
        {
            panel.add(btnArray.get(i), new GBC(0, 2 + i).setAnchor(GridBagConstraints.CENTER));
        }

        add(panel, new GBC(0, 0).setAnchor(GridBagConstraints.NORTHWEST).setInsets(5, 5, 0, 0));
    }

    public void initValue()
    {

    }

    private void initColumn()
    {
        column = new String[]{"№ п/п",
                "Назва товару",
                "Залишок на початок. Склад",
                "Залишок на початок. Магазин",
                "Прихід ABC",
                "Списано ABC",
                "Видано у Ro-Max",
                "Повернуто із Ro-Max",
                "Залишок на кінець. Склад",
                "Залишок на кінець. Магазин",
                "Різниця",
                "Ціна"};
        tableModel = new DefaultTableModel(new String[][]{}, column);
    }

    private void initTable()
    {
        ArrayList<String> arrayList = new ArrayList<>();
        ArrayList<String> row = new ArrayList<>();

        for (int i = 0; i < tableModel.getRowCount();)
            tableModel.removeRow(i);

        try
        {
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SHOW TABLES FROM `databassesabc`");
            while (resultSet.next())
            {
                try
                {
                    arrayList.add(main.getDataBaseDateFormat().format(main.getDataBaseDateFormat().parse(resultSet.getString(1))));
                } catch (ParseException ignored) {}
            }
        } catch (SQLException e) {e.printStackTrace();}

        if (!arrayList.isEmpty())
        {
            Collections.sort(arrayList);
            if (arrayList.contains(main.getDataBaseDateFormat().format(datePicker.getDate())))
            {
                String thisGroup = "";
                String group;
                try
                {
                    dbi = new DBI("databassesabc");
                    resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + ", `Група` FROM `" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "`;");
                    while (resultSet.next())
                    {
                        group = resultSet.getString(13);
                        if (!group.equals(thisGroup))
                        {
                            row.add("");
                            row.add("<html><b><font style=\"font-size:13pt;\">" + group + "</b></font></html>");
                            thisGroup = group;
                            tableModel.addRow(row.toArray());
                            row.clear();
                        }
                        for (int i = 1; i <= resultSet.getMetaData().getColumnCount(); i++)
                        {

                            if (i <= 2)
                                row.add(resultSet.getString(i));
                            else
                            {
                                if (i > 5 && i < 12)
                                    row.add("");
                                else
                                if (i != 13)
                                    row.add("" + resultSet.getDouble(i));
                            }
                        }
                        tableModel.addRow(row.toArray());
                        row.clear();
                    }
                    dbi.close(main.k);
                } catch (SQLException e) {e.printStackTrace();}
            }
        }
        setColumnsWidthColumnLabileTextWeight(table);
    }

    public void initNewRevision()
    {
        int n = JOptionPane.showConfirmDialog(main, "Ви впевнені, що хочете розпочати ревізію сьогодні?", "Попередження", JOptionPane.INFORMATION_MESSAGE);
        boolean b = true;

        try
        {
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT * FROM `" + main.getDataBaseDateFormat().format(new Date()) + "`;");
            if (resultSet.next())
                b = false;
            dbi.close(main.k);
        } catch (SQLException ignored) {}

        if (b)
        {
            if (n == JOptionPane.OK_OPTION)
            {
                try
                {
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("CREATE TABLE IF NOT EXISTS `" + main.getDataBaseDateFormat().format(new Date()) + "` (" +
                            " `№ п/п` int(11) NOT NULL AUTO_INCREMENT," +
                            " `Група` text," +
                            " `Назва товару` text," +
                            " `Залишок на початок. Склад` decimal(10,3) DEFAULT NULL," +
                            " `Залишок на початок. Магазин` decimal(10,3) DEFAULT NULL," +
                            " `Прихід ABC` decimal(10,3) DEFAULT NULL," +
                            " `Списано ABC` decimal(10,3) DEFAULT NULL," +
                            " `Видано у Ro-Max` decimal(10,3) DEFAULT NULL," +
                            " `Повернуто із Ro-Max` decimal(10,3) DEFAULT NULL," +
                            " `Залишок на кінець. Склад` decimal(10,3) DEFAULT NULL," +
                            " `Залишок на кінець. Магазин` decimal(10,3) DEFAULT NULL," +
                            " `Різниця` decimal(10,3) DEFAULT NULL," +
                            " `Ціна` decimal(10,3) DEFAULT NULL," +
                            " PRIMARY KEY (`№ п/п`)" +
                            ") ENGINE=InnoDB DEFAULT CHARSET=utf8 ROW_FORMAT=COMPACT;");
                    dbi.close(main.k);
                } catch (SQLException e) {e.printStackTrace();}

                getFromCategory("Дрібля");
                getFromCategory("Канцтовари");
                getFromCategory("Хімія");
                getFromCategory("Продукти харчування");
                getFromCategory("Іграшки");
               /*
                try
                {
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("UPDATE `товари` SET " +
                            "`Залишок на початок. Склад`='0', " +
                            "`Залишок на початок. Магазин`='0', " +
                            "`Прихід`='0';");
                    dbi.close(main.k);
                } catch (SQLException e) {e.printStackTrace();} */
                createExcelForCategory("Дрібля");
                createExcelForCategory("Канцтовари");
                createExcelForCategory("Хімія");
                createExcelForCategory("Продукти харчування");
                createExcelForCategory("Іграшки");
                initTable();
                n = JOptionPane.showConfirmDialog(main, "Формування бази для ревізії завершено.\nБажаєте відкрити сформовані накладні?", "Повідомлення", JOptionPane.YES_NO_OPTION);
                if (n == JOptionPane.OK_OPTION)
                {
                    for (String aCategory : category)
                        try
                        {
                            Runtime.getRuntime().exec("cmd /c start excel \"" + aCategory + "\"");
                        } catch (IOException e)
                        {
                            e.printStackTrace();
                        }
                }
            }
        } else
        {
            JOptionPane.showMessageDialog(main, "Неможливо проводити одночасно дві ревізії!");
        }
    }

    private void createExcelForCategory(String category)
    {
        HSSFWorkbook workbook = readWorkbook("Ревізія.xls");
        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        try
        {
            String thisGroup = "";
            String group;
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT `Група`, `Назва товару`, `Залишок на початок. Склад`, `Залишок на початок. Магазин`, `Прихід`, `Ціна` " +
                    "FROM `товари` " +
                    "WHERE `Категорія`='" + category + "' " +
                    "ORDER BY `Група` ASC;");
            for (int i = 2, k = 1; resultSet.next(); i++, k++)
            {
                row = sheet.getRow(i);
                group = resultSet.getString(1);
                if (!thisGroup.equals(group))
                {
                    i++;
                    cell = row.getCell(1);
                    cell.setCellStyle(sheet.getRow(0).getCell(0).getCellStyle());
                    cell.setCellValue(group);
                    thisGroup = group;
                    row = sheet.getRow(i);
                }
                row.getCell(0).setCellValue(k);
                row.getCell(1).setCellValue(resultSet.getString(2));
                row.getCell(2).setCellValue(resultSet.getDouble(3));
                row.getCell(3).setCellValue(resultSet.getDouble(4));
                row.getCell(4).setCellValue(resultSet.getDouble(5));
                row.getCell(10).setCellFormula("C" + (i + 1) + "+" +
                        "D" + (i + 1) + "+" +
                        "E" + (i + 1) + "-" +
                        "F" + (i + 1) + "-" +
                        "G" + (i + 1) + "+" +
                        "H" + (i + 1) + "-" +
                        "I" + (i + 1) + "-" +
                        "J" + (i + 1));
                row.getCell(11).setCellValue(resultSet.getDouble(6));
            }
        } catch (SQLException e) {e.printStackTrace();}
        try
        {
            FileOutputStream fileOut = new FileOutputStream(category + ".xls");
            workbook.write(fileOut);
            fileOut.close();
        } catch (IOException e1) {e1.printStackTrace();}
    }

    void getFromCategory(String category)
    {
        DBI dbiWarehouse;
        ResultSet warehouseResultSet;
        DBI dbiRevision;
        try
        {
            dbiWarehouse = new DBI("databassesabc");
            warehouseResultSet = dbiWarehouse.getSt().executeQuery("SELECT `Назва товару`, `Група`, `Залишок на початок. Склад`, `Залишок на початок. Магазин`, `Прихід`, `Ціна` " +
                    "FROM `товари` " +
                    "WHERE `Категорія`='" + category + "' " +
                    "ORDER BY `Група` ASC;");
            while (warehouseResultSet.next())
            {
                dbiRevision = new DBI("databassesabc");
                dbiRevision.getSt().execute("INSERT INTO `" + main.getDataBaseDateFormat().format(new Date()) + "` (" +
                        "`Назва товару`, " +
                        "`Група`, " +
                        "`Залишок на початок. Склад`, " +
                        "`Залишок на початок. Магазин`, " +
                        "`Прихід ABC`, " +
                        "`Ціна`) VALUES (" +
                        "'" + warehouseResultSet.getString(1) + "', " +
                        "'" + warehouseResultSet.getString(2) + "', " +
                        "'" + warehouseResultSet.getString(3) + "', " +
                        "'" + warehouseResultSet.getString(4) + "', " +
                        "'" + warehouseResultSet.getString(5) + "', " +
                        "'" + warehouseResultSet.getString(6) + "');");
                dbiRevision.close(main.k);
            }
        } catch (SQLException e) {e.printStackTrace();}
    }

    private class ImportBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e) {
            int n = JOptionPane.showConfirmDialog(main, "Ви впевнені, що бажаєте імпотрувати дані з документів Excel?", "Імпорт", JOptionPane.YES_NO_OPTION);
            if (n == JOptionPane.OK_OPTION)
            {
                for (String aCategory : category)
                {
                    HSSFWorkbook workbook = readWorkbook(aCategory + ".xls");
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    HSSFRow row;
                    HSSFCell cell;     /*
                    String name = null;
                    {
                        row = sheet.getRow(i);
                        name = row.getCell(1).getStringCellValue();
                        try
                        {

                        }
                    } while (!name.equals(""));    */
                }
            }
        }
    }
}
