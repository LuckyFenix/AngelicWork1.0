package Waybills;

import Main.Main;
import RComponents.RPanel;
import Support.DBI;
import Support.GBC;
import com.michaelbaranov.microba.calendar.DatePicker;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Row;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.*;

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
                "Залишок на початок. Ro-Max",
                "Прихід ABC",
                "Списано ABC",
                "Видано у Ro-Max",
                "Залишок у Ro-Max",
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
                        group = resultSet.getString(14);
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
                                if (i > 5 && i < 13)
                                    row.add("");
                                else
                                if (i != 14)
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
        } catch (SQLException e) {}

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
                            " `Залишок на початок. Ro-Max` decimal(10,3) DEFAULT NULL," +
                            " `Прихід ABC` decimal(10,3) DEFAULT NULL," +
                            " `Списано ABC` decimal(10,3) DEFAULT NULL," +
                            " `Видано у Ro-Max` decimal(10,3) DEFAULT NULL," +
                            " `Залишок у Ro-Max` decimal(10,3) DEFAULT NULL," +
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
        HSSFWorkbook workbook = readShablonWorkbook("Ревізія.xls");
        HSSFSheet sheet = workbook.getSheetAt(0);
        HSSFRow row;
        HSSFCell cell;
        try
        {
            String thisGroup = "";
            String group;
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT `Група`, `Назва товару`, `Одиниці вимірювання`, `Залишок на початок. Склад`, `Залишок на початок. Магазин`, `Прихід`, `Ціна` " +
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
                row.getCell(5).setCellValue(resultSet.getDouble(6));
                row.getCell(11).setCellFormula("C" + (i + 1) + "+" +
                        "D" + (i + 1) + "+" +
                        "E" + (i + 1) + "-" +
                        "F" + (i + 1) + "-" +
                        "G" + (i + 1) + "+" +
                        "H" + (i + 1) + "-" +
                        "I" + (i + 1) + "-" +
                        "J" + (i + 1));
                row.getCell(12).setCellValue(resultSet.getDouble(7));
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
                    "ORDER BY `Група`, `Назва товару` ASC;");
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

    public void saveRevision()
    {
        int n = JOptionPane.showConfirmDialog(main, "Ви впевнені, що бажаєте зберегти дані?", "Зберегти?", JOptionPane.YES_NO_OPTION);
        if (n == JOptionPane.OK_OPTION)
        {
            try
            {
                dbi = new DBI("databassesabc");
                dbi.getSt().executeUpdate("UPDATE `Товари` SET " +
                        "`Залишок на початок. Магазин` = '0', " +
                        "`Залишок на початок. Склад` = '0', " +
                        "`Товару в загальному` = '0', " +
                        "`Прихід` = '0';");
                dbi.close(main.k);
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT * FROM `" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "`;");
                while (resultSet.next())
                {
                    Map<String, String> map = new HashMap<>();
                    ResultSetMetaData metaData = resultSet.getMetaData();
                    for (int i = 0; i < resultSet.getMetaData().getColumnCount(); i++)
                    {
                        map.put(metaData.getColumnLabel(i+1), resultSet.getString(i+1));
                    }
                    DBI dbi1 = new DBI("databassesabc");
                    String query = "SELECT * FROM `товари` " +
                            "WHERE `Назва товару` = '" + map.get("Назва товару") + "' " +
                            "AND `Ціна` = '" + map.get("Ціна") + "';";
                    System.out.println(query);
                    ResultSet resultSet1 = dbi1.getSt().executeQuery(query);
                    if (resultSet1.next())
                    {
                        DBI dbi2 = new DBI("databassesabc");
                        query = "UPDATE `товари` " +
                                "SET " +
                                "`Залишок на початок. Склад` = '" + map.get("Залишок на кінець. Склад") + "', " +
                                "`Залишок на початок. Магазин` = '" + map.get("Залишок на кінець. Магазин") + "', " +
                                "`Прихід` = '0', " +
                                "`Товару в загальному` = '" + (Double.parseDouble(map.get("Залишок на кінець. Склад")) + Double.parseDouble(map.get("Залишок на кінець. Магазин"))) + "' " +
                                "WHERE " +
                                "`Назва товару` = '" + map.get("Назва товару") + "' " +
                                "AND " +
                                "`Ціна` = '" + map.get("Ціна") + "';";
                        System.out.println(query);
                        dbi2.getSt().executeUpdate(query);
                        dbi2.close(main.k);
                    } else
                    {
                        String category;
                        DBI dbi2 = new DBI("databassesabc");
                        ResultSet resultSet2 = dbi2.getSt().executeQuery("SELECT `Категорія` FROM `товари` " +
                                "WHERE `Група` = '" + map.get("Група") + "';");
                        if (resultSet2.next())
                        {
                            category = resultSet2.getString("Категорія");
                            resultSet2.close();
                            dbi2.close(main.k);
                            dbi2 = new DBI("databassesabc");
                            query = "INSERT INTO `товари` (" +
                                    "`Категорія`, " +
                                    "`Група`, " +
                                    "`Назва товару`, " +
                                    "`Одиниці вимірювання`, " +
                                    "`Залишок на початок. Склад`, " +
                                    "`Залишок на початок. Магазин`, " +
                                    "`Товару в загальному`, " +
                                    "`Ціна`, " +
                                    "`Ціна закупочна`, " +
                                    "`Дата останнього приходу`) " +
                                    "VALUES (" +
                                    "'" + category + "', " +
                                    "'" + map.get("Група") + "', " +
                                    "'" + map.get("Назва товару") + "', " +
                                    "'шт.', " +
                                    "'" + map.get("Залишок на кінець. Склад") + "', " +
                                    "'" + map.get("Залишок на кінець. Магазин") + "', " +
                                    "'" + (Double.parseDouble(map.get("Залишок на кінець. Склад")) + Double.parseDouble(map.get("Залишок на кінець. Магазин"))) + "', " +
                                    "'" + map.get("Ціна") + "', " +
                                    "'0', " +
                                    "'0000-00-00');";
                            System.out.println(query);
                            dbi2.getSt().executeUpdate(query);
                            dbi2.close(main.k);
                        }
                        resultSet1.close();
                        dbi1.close(main.k);
                    }
                }
                resultSet.close();
                dbi.close(main.k);
                JOptionPane.showMessageDialog(main, "Сохранение успешно завершено");
            } catch (Exception e1)
            {
                JOptionPane.showMessageDialog(main, e1, "Ошибка", JOptionPane.ERROR_MESSAGE);
                e1.printStackTrace();
            }
        }
    }

    public void nullDelete()
    {
        String date = main.getDataBaseDateFormat().format(datePicker.getDate());
        try
        {
            ArrayList names = new ArrayList();
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT " +
                    "`товари`.`Назва товару` " +
                    "FROM " +
                    "`товари`, " +
                    "`" + date + "` " +
                    "WHERE " +
                    "`товари`.`Назва товару`=`" + date + "`.`Назва товару` " +
                    "AND " +
                    "`товари`.`Ціна`=`" + date + "`.`Ціна` " +
                    "AND" +
                    "`товари`.`Залишок на початок. Магазин`='0'" +
                    "AND" +
                    "`товари`.`Залишок на початок. Склад`='0'" +
                    "AND" +
                    "`" + date + "`.`Залишок у Ro-Max`='0' " +
                    "ORDER BY `товари`.`Категорія`, `товари`.`Група`, `товари`.`Назва товару` ASC;");
            while (resultSet.next())
            {
                names.add(resultSet.getString("Назва товару"));
            }
            resultSet.close();
            dbi.close(main.k);
            String[] column = {"Група", "Назва товару", "Магазин", "Склад", "Ro-Max", "Ціна", "Ціна прихідна", "Дата останнього приходу"};
            DefaultTableModel tableModel = new DefaultTableModel(column, 0);
            for (int i = 0; i < names.size(); i++)
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT " +
                        "`товари`.`Група`, " +
                        "`товари`.`Назва товару`, " +
                        "`товари`.`Залишок на початок. Магазин`, " +
                        "`товари`.`Залишок на початок. Склад`, " +
                        "`2014-07-01`.`Залишок у Ro-Max`, " +
                        "`товари`.`Ціна`, " +
                        "`товари`.`Ціна закупочна`, " +
                        "`товари`.`Дата останнього приходу` " +
                        "FROM " +
                        "`товари`, " +
                        "`2014-07-01` " +
                        "WHERE " +
                        "`товари`.`Назва товару`=`2014-07-01`.`Назва товару` " +
                        "AND " +
                        "`товари`.`Ціна`=`2014-07-01`.`Ціна` " +
                        "AND " +
                        "`товари`.`Назва товару`='" + names.get(i) + "';");
                while (resultSet.next())
                {
                    String d;
                    try
                    {
                        d = resultSet.getString(8);
                    } catch (SQLException e) {d = "0000-00-00";}

                    tableModel.addRow(new Object[]
                            {
                                    resultSet.getString(1),
                                    resultSet.getString(2),
                                    resultSet.getString(3),
                                    resultSet.getString(4),
                                    resultSet.getString(5),
                                    resultSet.getString(6),
                                    resultSet.getString(7),
                                    d
                            });
                }
                resultSet.close();
                dbi.close(main.k);
            }
            JTable table = new JTable(tableModel);
            table.addMouseListener(new MouseAdapter()
            {
                @Override
                public void mouseReleased(MouseEvent e)
                {
                    super.mouseReleased(e);
                    if (e.getButton() == MouseEvent.BUTTON3)
                    {
                        int n = JOptionPane.showConfirmDialog(main, "Удалить");
                        if (n == JOptionPane.OK_OPTION)
                        {
                            try
                            {
                                dbi = new DBI("databassesabc");
                                dbi.getSt().execute("DELETE FROM `Товари` " +
                                        "WHERE `Назва товару`='" + table.getValueAt(table.getSelectedRow(), 1) + "' " +
                                        "AND `Ціна`='" + table.getValueAt(table.getSelectedRow(), 5) + "';");
                                dbi.close(main.k);
                                tableModel.removeRow(table.getSelectedRow());
                            } catch (SQLException e1)
                            {
                                e1.printStackTrace();
                            }
                        }
                    }
                }
            });
            setColumnsWidthColumnLabileTextWeight(table);
            JScrollPane scrollPane = new JScrollPane(table);
            scrollPane.setMinimumSize(new Dimension(700, 1000));
            JOptionPane.showMessageDialog(main, scrollPane);

        } catch (SQLException e)
        {
            e.printStackTrace();
        }
    }

    public void exportToExcel()
    {
        String date = main.getDataBaseDateFormat().format(datePicker.getDate());
        int n = JOptionPane.showConfirmDialog(main, "Ви впевнені, що бажаєте зберегти дані?", "Зберегти?", JOptionPane.YES_NO_OPTION);
        if (n == JOptionPane.OK_OPTION)
        {
            try
            {
                dbi = new DBI("databassesabc");
//                resultSet = dbi.getSt().executeQuery("SELECT `товари`.*, `" + date + "`.`Залишок у Ro-Max` " +
//                        "FROM `товари`, `" + date + "` " +
//                        "WHERE `товари`.`Назва товару`=`" + date + "`.`Назва товару` " +
//                        "AND `товари`.`Ціна`=`" + date + "`.`Ціна`" +
//                        "ORDER BY `Категорія`, `Група`, `Назва товару` ASC;");
                resultSet = dbi.getSt().executeQuery("SELECT * FROM `товари` " +
                        "ORDER BY `Категорія`, `Група`, `Назва товару` ASC;");
                HSSFWorkbook workbook = readShablonWorkbook("Ревізія реімпорт.xls");
                HSSFSheet sheet = workbook.getSheetAt(0);
                for (int i = 2; resultSet.next(); i++)
                {
                    String name = resultSet.getString("Назва товару");
                    double price = resultSet.getDouble("Ціна");
                    DBI dbi1 = new DBI("databassesabc");
                    ResultSet resultSet1 = dbi1.getSt().executeQuery("SELECT `Залишок у Ro-Max` FROM `" + date + "` " +
                            "WHERE `Ціна`='" + price + "' " +
                            "AND `Назва товару`='" + name + "';");
                    resultSet1.next();
                    HSSFRow row = sheet.getRow(i);
                    row.getCell(1).setCellValue(resultSet.getString("Категорія"));
                    row.getCell(2).setCellValue(resultSet.getString("Група"));
                    row.getCell(3).setCellValue(resultSet.getString("Назва товару"));
                    row.getCell(4).setCellValue(resultSet.getString("Одиниці вимірювання"));
                    row.getCell(5).setCellValue(resultSet.getString("Залишок на початок. Склад"));
                    row.getCell(6).setCellValue(resultSet.getString("Залишок на початок. Магазин"));
                    row.getCell(7).setCellValue(resultSet1.getString("Залишок у Ro-Max"));
                    row.getCell(8).setCellValue(resultSet.getString("Ціна"));
                    row.getCell(9).setCellValue(resultSet.getString("Ціна закупочна"));
                    resultSet1.close();
                    dbi1.close(main.k);
                }
                FileOutputStream fileOut = new FileOutputStream("Ревізія реімпорт.xls");
                workbook.write(fileOut);
                fileOut.close();
                JOptionPane.showMessageDialog(main, "Експорт завершено");
            } catch (Exception e) {e.printStackTrace();}
        }
    }

    private class ImportBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            int n = JOptionPane.showConfirmDialog(main, "Ви впевнені, що бажаєте імпотрувати дані з документів Excel?", "Імпорт", JOptionPane.YES_NO_OPTION);
            if (n == JOptionPane.OK_OPTION)
            {
                for (String aCategory : category)
                {
                    String group = "";
                    String query = "";
                    boolean end = false;
                    HSSFWorkbook workbook = readWorkbook(aCategory + ".xls");
                    HSSFSheet sheet = workbook.getSheetAt(0);
                    Row row;
                    HSSFCell cell;
                    Iterator<Row> iterator = sheet.rowIterator();
                    iterator.next();
                    row = iterator.next();
                    while (iterator.hasNext() || !end)
                    {
                        try
                        {
                            if (row.getCell(12).getNumericCellValue() == 0)
                            {
                                group = row.getCell(1).getStringCellValue();
                            } else
                            {
                                if (row.getCell(0).getNumericCellValue() == 0)
                                {
                                    String cell1 = row.getCell(1).getStringCellValue();
                                    double cell2 = row.getCell(2).getNumericCellValue();
                                    double cell3 = row.getCell(3).getNumericCellValue();
                                    double cell4 = row.getCell(4).getNumericCellValue();
                                    double cell5 = row.getCell(5).getNumericCellValue();
                                    double cell6 = row.getCell(6).getNumericCellValue();
                                    double cell7 = row.getCell(7).getNumericCellValue();
                                    double cell8 = row.getCell(8).getNumericCellValue();
                                    double cell9 = row.getCell(9).getNumericCellValue();
                                    double cell10 = row.getCell(10).getNumericCellValue();
                                    double cell11 = row.getCell(11).getNumericCellValue();
                                    double cell12 = row.getCell(12).getNumericCellValue();
                                    try
                                    {
                                        dbi = new DBI("databassesabc");
                                        query = "INSERT INTO `" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "` (" +
                                                "`Група`, " +
                                                "`" + column[1] + "`, " +
                                                "`" + column[2] + "`, " +
                                                "`" + column[3] + "`, " +
                                                "`" + column[4] + "`, " +
                                                "`" + column[5] + "`, " +
                                                "`" + column[6] + "`, " +
                                                "`" + column[7] + "`, " +
                                                "`" + column[8] + "`, " +
                                                "`" + column[9] + "`, " +
                                                "`" + column[10] + "`, " +
                                                "`" + column[11] + "`, " +
                                                "`" + column[12] + "`) " +
                                                "VALUES (" +
                                                "'" +  group + "', " +
                                                "'" + cell1 + "', " +
                                                "'" + cell2 + "', " +
                                                "'" + cell3 + "', " +
                                                "'" + cell4 + "', " +
                                                "'" + cell5 + "', " +
                                                "'" + cell6 + "', " +
                                                "'" + cell7 + "', " +
                                                "'" + cell8 + "', " +
                                                "'" + cell9 + "', " +
                                                "'" + cell10 + "', " +
                                                "'" + cell11 + "', " +
                                                "'" + cell12 + "');";
                                        System.out.println(query);
                                        dbi.getSt().executeUpdate(query);
                                        dbi.close(main.k);
                                    } catch (SQLException e1)
                                    {
                                        e1.printStackTrace();
                                    }
                                } else
                                {
                                    try
                                    {
                                        dbi = new DBI("databassesabc");
                                        query = "UPDATE `" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "` " +
                                                "SET " +
                                                "`" + column[4] + "` = '" + row.getCell(4).getNumericCellValue() + "', " +
                                                "`" + column[6] + "` = '" + row.getCell(6).getNumericCellValue() + "', " +
                                                "`" + column[7] + "` = '" + row.getCell(7).getNumericCellValue() + "', " +
                                                "`" + column[8] + "` = '" + row.getCell(8).getNumericCellValue() + "', " +
                                                "`" + column[9] + "` = '" + row.getCell(9).getNumericCellValue() + "', " +
                                                "`" + column[10] + "` = '" + row.getCell(10).getNumericCellValue() + "', " +
                                                "`" + column[11] + "` = '" + row.getCell(11).getNumericCellValue() + "' " +
                                                "WHERE " +
                                                "`" + column[1] + "` = '" + row.getCell(1).getStringCellValue() + "' " +
                                                "AND " +
                                                "`" + column[12] + "` = '" + row.getCell(12).getNumericCellValue() + "';";
                                        System.out.println(query);
                                        dbi.getSt().execute(query);
                                        dbi.close(main.k);
                                    } catch (SQLException e1) {e1.printStackTrace();}
                                }
                            }
                        } catch (Exception e1) {e1.printStackTrace();}
                        row = iterator.next();
                        if (row.getCell(1).getStringCellValue().trim().equals(""))
                            end = true;
                    }
                }
                JOptionPane.showMessageDialog(main, "Імпорт завершено");
            }
        }
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

    public HSSFWorkbook readShablonWorkbook(String filename)
    {
        try
        {
            POIFSFileSystem fs = new POIFSFileSystem(getClass().getResourceAsStream("/" + filename));
            return new HSSFWorkbook(fs);
        }
        catch (Exception e) {
            return null;
        }
    }
}
