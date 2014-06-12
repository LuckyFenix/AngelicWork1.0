package Waybills;

import Main.Main;
import RComponents.RPanel;
import RDialogs.AddWaybillDialog;
import RDialogs.RDialog;
import Support.DBI;
import Support.GBC;
import com.michaelbaranov.microba.calendar.DatePicker;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.beans.PropertyVetoException;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;

public class WaybillSelling extends RPanel
{
    private final Main main;
    private JPanel panel = new JPanel(new GridBagLayout());
    private DatePicker datePicker;
    private JTextField staffField;
    private JTextField noField;
    private DefaultTableModel tableModel;
    private JTable table;
    private String[] column;
    private ArrayList<String> waybills = new ArrayList<>();
    private DBI dbi;
    private ResultSet resultSet;
    private Date revisionDate;

    public WaybillSelling(Main main)
    {
        this.main = main;

        setLayout(new GridBagLayout());

        initPanel();
    }

    private void initPanel()
    {
        this.removeAll();

        noField = new JTextField(10);
        datePicker = new DatePicker();
        staffField = new JTextField(15);
        JButton staffBtn = new JButton("...");
        JButton addBtn = new JButton("Додати накладні");
        JButton exportBtn = new JButton("Експортувати в Excel", new ImageIcon(getClass().getResource("/icon/excel_icon.png")));
        JButton exportForReimportBtn = new JButton("Експортувати в Excel для реімпорту", new ImageIcon("Resourse/icon/excel_icon.png"));
        tableModel = new DefaultTableModel();
        table = new JTable(tableModel);

        column = new String[]{"Артикль", "Назва товару", "Одиниці вимірювання", "Кількість", "Ціна", "Сума"};
        tableModel.setColumnIdentifiers(column);

        noField.setEditable(false);
        datePicker.setFieldEditable(false);

        datePicker.addActionListener(e ->
        {
            initNo();
            initTable();
        });
        staffBtn.addActionListener(e -> new AddStaffDialog());
        addBtn.addActionListener(e -> new AddWaybillDialog(noField.getText(), datePicker.getDate(), staffField.getText(), main));
        exportBtn.addActionListener(new ExportBtnListener());
        exportForReimportBtn.addActionListener(new ExportForReimportBtnListener());

        initNo();
        initTable();

        JScrollPane scrollPane = new JScrollPane(table);

        panel.add(new JLabel("Номер накладної:"), new GBC(0, 0).setAnchor(GridBagConstraints.WEST).setInsets(0, 0, 0, 5));
        panel.add(noField, new GBC(1, 0, 2, 1).setAnchor(GridBagConstraints.EAST));
        panel.add(new JLabel("Дата:"), new GBC(0, 1).setAnchor(GridBagConstraints.WEST));
        panel.add(datePicker, new GBC(1, 1, 2, 1).setAnchor(GridBagConstraints.EAST));
        panel.add(new JLabel("Відпущено кому:"), new GBC(0, 2).setAnchor(GridBagConstraints.WEST));
        panel.add(staffField, new GBC(1, 2).setAnchor(GridBagConstraints.EAST));
        panel.add(staffBtn, new GBC(2, 2).setAnchor(GridBagConstraints.EAST));
        panel.add(addBtn, new GBC(0, 3, 3, 1).setAnchor(GridBagConstraints.EAST));
        panel.add(exportBtn, new GBC(0, 4, 3, 1).setAnchor(GridBagConstraints.EAST));
        panel.add(exportForReimportBtn, new GBC(0, 5, 3, 1).setAnchor(GridBagConstraints.NORTHEAST));
        panel.add(scrollPane, new GBC(3, 0, 1, 7).setIpad(scrollPane.getWidth() + 300, scrollPane.getHeight()));

        add(panel, new GBC(0, 0).setAnchor(GridBagConstraints.NORTHWEST).setInsets(10));
    }

    private void initNo()
    {
        ArrayList<String> arrayList = new ArrayList<>();
        try {
            revisionDate = main.getDataBaseDateFormat().parse("0000-00-00");
        } catch (ParseException e) {e.printStackTrace();}

        try
        {
            DBI dbi = new DBI("databassesabc");
            ResultSet resultSet = dbi.getSt().executeQuery("SHOW TABLES FROM `databassesabc`");
            while (resultSet.next())
            {
                try
                {
                    arrayList.add(main.getDataBaseDateFormat().format(main.getDataBaseDateFormat().parse(resultSet.getString(1))));
                } catch (ParseException ignored) {}
            }
            if (!arrayList.isEmpty())
            {
                Collections.sort(arrayList);
                revisionDate = main.getDataBaseDateFormat().parse(arrayList.get(arrayList.size() - 1));
            }
        } catch(SQLException e) {e.printStackTrace();} catch (ParseException e) {e.printStackTrace();}

        try
        {
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT `Дата відпуску` FROM `відпускні накладні` " +
                    "WHERE `Дата відпуску`>'" + main.getDataBaseDateFormat().format(revisionDate) + "' " +
                    "GROUP BY `Номер відпускної накладної`;");
            noField.setText("" + (getResultSetSize(resultSet)));
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}
    }

    public void initTable()
    {
        waybills.clear();

        for(int i = 0; i < tableModel.getRowCount();)
        {
            tableModel.removeRow(i);
        }

        try
        {
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT `Дата`, `Номер` FROM `відпускні накладні` " +
                    "WHERE `Номер відпускної накладної`='" + noField.getText().trim() + "'" +
                    "AND `Дата відпуску`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "';");
            while (resultSet.next())
            {
                waybills.add(main.getDataBaseDateFormat().format(resultSet.getDate(1)) + "@" + resultSet.getString(2));
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        try
        {
            ArrayList<String> row = new ArrayList<>();
            for (String waybill : waybills)
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `прихід` " +
                        "WHERE `Дата`='" + waybill.split("@")[0] + "'" +
                        "AND `Номер`='" + waybill.split("@")[1] + "';");
                while (resultSet.next())
                {
                    for (int j = 1; j <= resultSet.getMetaData().getColumnCount(); j++)
                    {
                        if (row.size() == resultSet.getMetaData().getColumnCount() - 1)
                        {
                            double d = Double.parseDouble(row.get(row.size() - 1)) * Double.parseDouble(row.get(row.size() - 2));
                            d = d * 1000;
                            int integ = (int) Math.round(d);
                            d = (double) integ / 1000;
                            row.add("" + d);
                        } else
                            row.add(resultSet.getString(j));
                    }
                    tableModel.addRow(row.toArray());
                    row.clear();
                }
                dbi.close(main.k);
            }
        } catch (SQLException e) {e.printStackTrace();}

        setColumnsWidth(table);
    }

    public void initTable(Date date, String no, String staff)
    {
        try {
            datePicker.setDate(date);
        } catch (PropertyVetoException e) {e.printStackTrace();}
        noField.setText(no);
        staffField.setText(staff);
        initTable();
    }

    private class AddStaffDialog extends RDialog
    {
        public AddStaffDialog()
        {
            super(main, "Працівники", true);

            GridBagLayout layout = new GridBagLayout();
            setLayout(layout);

            String[] s = new String[0];

            try
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT `Імя` FROM `працівники`;");
                s = new String[getResultSetSize(resultSet)];
                for (int i = 0; resultSet.next(); i++)
                {
                    s[i] = resultSet.getString(1);
                }
            } catch (SQLException e) {e.printStackTrace();}

            final JList<String> list = new JList<>(s);
            JButton addStaffBtn = new JButton("Додати");
            JButton closeBtb = new JButton("Закрити");

            addStaffBtn.addActionListener(e ->
            {
                if (staffField.getText().equals(""))
                {
                    staffField.setText(list.getSelectedValue());
                }
                else
                    staffField.setText(staffField.getText() + ", " + list.getSelectedValue());
                dispose();
            });

            closeBtb.addActionListener(e -> dispose());

            list.addMouseListener(new MouseAdapter()
            {
                @Override
                public void mouseClicked(MouseEvent e)
                {
                    if (e.getClickCount() == 2)
                    {
                        if (staffField.getText().equals(""))
                        {
                            staffField.setText(list.getSelectedValue());
                        }
                        else
                            staffField.setText(staffField.getText() + ", " + list.getSelectedValue());
                        dispose();
                    }
                }
            });

            add(list, new GBC(0, 0, 3, 1).setAnchor(GridBagConstraints.NORTHWEST).setIpad(200, 100));
            add(addStaffBtn, new GBC(1, 1).setAnchor(GridBagConstraints.EAST));
            add(closeBtb, new GBC(2, 1).setAnchor(GridBagConstraints.EAST));

            pack();

            setResizable(false);
            setLocationRelativeTo(main);
            setLocation(main.getX() + main.getWidth() / 2 - this.getWidth() / 2, main.getY() + main.getHeight() / 2 - this.getHeight() / 2);
            revalidate();
            setVisible(true);
        }
    }

    private class ExportBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            HSSFWorkbook workbook = readWorkbook("Накладні.xls");
            HSSFSheet sheet = workbook.getSheetAt(1);
            HSSFRow row;
            HSSFCell cell;

            for (int i = 6, i1 = 0; i1 < table.getRowCount(); i++, i1++)
            {
                row = sheet.getRow(i);
                for (int j = 1, j1 = 1; j1 < table.getColumnCount(); j++, j1++)
                {
                    cell = row.getCell(j + 9);
                    if (j == 1 || j == 2)
                        cell.setCellValue(table.getValueAt(i1, j1).toString());
                    else
                    {
                        if (j == 5)
                            cell.setCellValue(Double.parseDouble(table.getValueAt(i1, j1 - 2).toString()) * Double.parseDouble(table.getValueAt(i1, j1 - 1).toString()));
                        else
                            cell.setCellValue(Double.parseDouble(table.getValueAt(i1, j1).toString()));
                    }
                    cell = row.getCell(j);
                    if (j == 1 || j == 2)
                    {
                        cell.setCellValue(table.getValueAt(i1, j1).toString());
                    }
                    else
                    {
                        if (j == 5)
                            cell.setCellValue(Double.parseDouble(table.getValueAt(i1, j1 - 2).toString()) * Double.parseDouble(table.getValueAt(i1, j1 - 1).toString()));
                        else
                            cell.setCellValue(Double.parseDouble(table.getValueAt(i1, j1).toString()));
                    }
                }

                if (i == 48 || i == 102)
                    i += 10;
                if (i == 156 || i == 209)
                    i += 9;
            }

            setFormulaToSheet(sheet, 49, 5, "SUM(F7:F49)");
            setFormulaToSheet(sheet, 49, 14, "SUM(O7:O49)");
            setFormulaToSheet(sheet, 103, 5, "SUM(F60:F103)");
            setFormulaToSheet(sheet, 103, 14, "SUM(O60:O103)");
            setFormulaToSheet(sheet, 157, 5, "SUM(F114:F157)");
            setFormulaToSheet(sheet, 157, 14, "SUM(O114:O157)");
            setFormulaToSheet(sheet, 210, 5, "SUM(F167:F210)");
            setFormulaToSheet(sheet, 210, 14, "SUM(O167:O210)");
            setFormulaToSheet(sheet, 263, 5, "SUM(F220:F263)");
            setFormulaToSheet(sheet, 263, 14, "SUM(O220:O263)");

            setFormulaToSheet(sheet, 104, 5, "F104+F50");
            setFormulaToSheet(sheet, 104, 14, "F104+F50");
            setFormulaToSheet(sheet, 158, 5, "F104+F50+F158");
            setFormulaToSheet(sheet, 158, 14, "F104+F50+F158");
            setFormulaToSheet(sheet, 211, 5, "F104+F50+F158+F211");
            setFormulaToSheet(sheet, 211, 14, "F104+F50+F158+F211");
            setFormulaToSheet(sheet, 264, 5, "F104+F50+F158+F211+F264");
            setFormulaToSheet(sheet, 264, 14, "F104+F50+F158+F211+F264");

            setValueToSheet(sheet, 0, 2, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 0, 11, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 53, 2, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 53, 11, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 107, 2, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 107, 11, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 160, 2, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 160, 11, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 213, 2, "Накладна № " + noField.getText());
            setValueToSheet(sheet, 213, 11, "Накладна № " + noField.getText());

            setValueToSheet(sheet, 2, 2, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 2, 11, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 55, 2, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 55, 11, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 109, 2, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 109, 11, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 162, 2, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 162, 11, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 215, 2, main.getDateFormat().format(datePicker.getDate()));
            setValueToSheet(sheet, 215, 11, main.getDateFormat().format(datePicker.getDate()));

            setValueToSheet(sheet, 3, 1, staffField.getText());
            setValueToSheet(sheet, 3, 10, staffField.getText());
            setValueToSheet(sheet, 56, 1, staffField.getText());
            setValueToSheet(sheet, 56, 10, staffField.getText());
            setValueToSheet(sheet, 110, 1, staffField.getText());
            setValueToSheet(sheet, 110, 10, staffField.getText());
            setValueToSheet(sheet, 163, 1, staffField.getText());
            setValueToSheet(sheet, 163, 10, staffField.getText());
            setValueToSheet(sheet, 216, 1, staffField.getText());
            setValueToSheet(sheet, 216, 10, staffField.getText());

            try
            {
                FileOutputStream fileOut = new FileOutputStream("new_workbook.xls");
                workbook.write(fileOut);
                fileOut.close();
            } catch (IOException e1) {e1.printStackTrace();}

            try
            {
                Runtime.getRuntime().exec("cmd /c start excel \"new_workbook.xls\"");
            } catch (IOException ignored) {}
        }
    }

    private class ExportForReimportBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            HSSFWorkbook workbook = readWorkbook("Реімпорт.xls");
            HSSFSheet sheet = workbook.getSheetAt(0);
            HSSFRow row;
            HSSFCell cell;

            ArrayList<Integer> noArray = new ArrayList<>();
            ArrayList<Date> dateArray = new ArrayList<>();

            try
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT `Дата`, `Номер` " +
                        "FROM `відпускні накладні` " +
                        "WHERE `Номер відпускної накладної`='" + noField.getText().trim() + "' " +
                        "AND `Дата відпуску`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "';");
                while (resultSet.next())
                {
                    noArray.add(resultSet.getInt(2));
                    dateArray.add(resultSet.getDate(1));
                }
                dbi.close(main.k);
                resultSet.close();
            } catch (SQLException e1) {e1.printStackTrace();}

            int n = 0;
            for (int i = 0; i < noArray.size(); i++)
            {
                try {
                    dbi = new DBI("databassesabc");
                    resultSet = dbi.getSt().executeQuery("SELECT `Назва товару`, `Одиниці вимірювання`, `Ціна закупочна`, `Кількість`, `Ціна` " +
                            "FROM `прихід` " +
                            "WHERE `Номер`='" + noArray.get(i) + "' " +
                            "AND `Дата`='" + main.getDataBaseDateFormat().format(dateArray.get(i)) + "';");
                    while (resultSet.next())
                    {
                        n++;
                        row = sheet.getRow(n);
                        for (int j = 0; j < 5; j++)
                        {
                            cell = row.getCell(j);
                            cell.setCellValue(resultSet.getString(j + 1));
                        }
                    }
                    dbi.close(main.k);
                    resultSet.close();
                } catch (SQLException e1)
                {
                    e1.printStackTrace();
                }
            }

            try
            {
                FileOutputStream fileOut = new FileOutputStream("Накладна для реімпорту.xls");
                workbook.write(fileOut);
                fileOut.close();
            } catch (FileNotFoundException e1) {e1.printStackTrace();} catch (IOException e1) {e1.printStackTrace();}

            try
            {
                Runtime.getRuntime().exec("cmd /c start excel \"Накладна для реімпорту.xls\"");
            } catch (IOException ignored) {}
        }
    }
}
