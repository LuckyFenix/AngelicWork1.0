package Waybills;

import RComponents.*;
import RDialogs.InvoiceNewNomenclatureDialog;
import RDialogs.RDialog;
import Support.DBI;
import Main.Main;
import Support.GBC;
import com.michaelbaranov.microba.calendar.DatePicker;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.MouseAdapter;
import java.awt.event.MouseEvent;
import java.beans.PropertyVetoException;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;

public class Invoice extends RPanel
{
    private ArrayList<String> supplierArray = new ArrayList<>();
    private DBI dbi;
    private DatePicker datePicker;
    private Main main;
    private JTextField noField;
    private JTextField supplier;
    private JTable table;
    private DefaultTableModel tableModel;
    private JScrollPane scrollPane;
    private String[] column;
    private String[][] tableFields;
    private ResultSet resultSet;
    private JLabel fullSumValue;

    public Invoice(Main main)
    {
        this.main = main;

        setLayout(new GridBagLayout());

        initPanel();
    }

    public void initPanel()
    {
        this.removeAll();

        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery("SELECT * FROM `постачальники`;");
            while (resultSet.next())
            {
                supplierArray.add(resultSet.getString(1));
            }
            resultSet.close();
        } catch (SQLException e) {e.printStackTrace();}
        dbi.close(main.k);

        noField = new JTextField(3);
        datePicker = new DatePicker();
        supplier = new JTextField(10);
        JButton supplierBtn = new JButton("...");
        JButton newNomenclature = new JButton("Нова номенклатура");

        datePicker.setFieldEditable(false);
        supplier.setEditable(false);
        initNo();

        initColumn();
        tableModel = new DefaultTableModel((new String[0][0]), column);
        table = new JTable(tableModel);
        initTable();
        tableModel.addTableModelListener(e -> fullSumValue.setText("" + getSum()));

        fullSumValue = new JLabel("" + getSum());

        initPopupMenu();

        supplierBtn.addActionListener(e -> new SelectSupplierDialog());
        datePicker.addActionListener(e -> {
            initNo();
            initTable();
        });
        newNomenclature.addActionListener(e -> new InvoiceNewNomenclatureDialog(main,
                noField.getText(),
                supplier.getText(),
                datePicker.getDate()));

        scrollPane = new JScrollPane(table);
        JPanel valuePanel = new JPanel(new GridBagLayout());

        add(new JLabel("№"), new GBC(0, 0).setAnchor(GBC.WEST));
        add(noField, new GBC(1, 0, 2, 1).setAnchor(GBC.WEST));
        add(new JLabel("Дата"), new GBC(0, 1).setAnchor(GBC.WEST));
        add(datePicker, new GBC(1, 1, 2, 1).setAnchor(GBC.WEST));
        add(new JLabel("Постачальник"), new GBC(0, 2).setAnchor(GBC.WEST));
        add(supplier, new GBC(1, 2).setAnchor(GBC.WEST));
        add(supplierBtn, new GBC(2, 2).setAnchor(GridBagConstraints.WEST));
        add(newNomenclature, new GBC(0, 3, 3, 1).setAnchor(GBC.NORTH));
        add(scrollPane, new GBC(3, 0, 1, 4).setIpad(scrollPane.getWidth() + 300, scrollPane.getHeight() + 200));

        valuePanel.add(new JLabel("Загальна сума:    "), new GBC(0, 0));
        valuePanel.add(fullSumValue, new GBC(1, 0));

        add(valuePanel, new GBC(3, 4).setAnchor(GridBagConstraints.NORTHEAST));
    }

    public double getSum()
    {
        double s = 0;
        for (int i = 0; i < tableModel.getRowCount(); i++)
        {
            s += Double.parseDouble(table.getValueAt(i, 6).toString());
        }
        return s;
    }


    private void initColumn()
    {
        column = new String[]{"Артикль", "Назва товару", "Одиниці вимірювання", "Ціна закупочна", "Кількість", "Ціна", "Сума"};
    }

    void initNo()
    {
        dbi = new DBI("databassesabc");
        int max = 0;
        try
        {
            ResultSet resultSet = dbi.getSt().executeQuery("SELECT `Номер` FROM `прихід` WHERE `Дата`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "';");
            while (resultSet.next())
            {
                int i = Integer.parseInt(resultSet.getString(1));
                if(i > max)
                {
                    max = i;
                }
            }
            resultSet.close();
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        noField.setText("" + (max + 1));
        noField.setEditable(false);
    }

    public void initTable()
    {
        tableFields = new String[0][0];
        ResultSet resultSet;

        for(int i = 0; i < tableModel.getRowCount();)
        {
            tableModel.removeRow(i);
        }

        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " " +
                    "FROM `прихід` " +
                    "WHERE `Дата`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "' " +
                    "AND `Номер`='" + noField.getText() + "';");
            tableFields = new String[getResultSetSize(resultSet) - 1][resultSet.getMetaData().getColumnCount()];
            for (int i = 0; resultSet.next(); i++)
            {
                for (int j = 0; j < tableFields[0].length; j++)
                {
                    tableFields[i][j] = resultSet.getString(j + 1);
                }
                tableModel.addRow(tableFields[i]);
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        setColumnsWidth(table);
    }

    public void initTable(Date date, String no, String supplier)
    {
        try
        {
            datePicker.setDate(date);
            noField.setText(no);
            this.supplier.setText(supplier);
        } catch (PropertyVetoException e) {
            e.printStackTrace();
        }

        tableFields = new String[0][0];
        ResultSet resultSet;

        for(int i = 0; i < tableModel.getRowCount();)
        {
            tableModel.removeRow(i);
        }

        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `прихід` " +
                    "WHERE `Дата`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "' " +
                    "AND `Номер`='" + noField.getText() + "';");
            tableFields = new String[getResultSetSize(resultSet) - 1][resultSet.getMetaData().getColumnCount()];
            for (int i = 0; resultSet.next(); i++)
            {
                for (int j = 0; j < tableFields[0].length; j++)
                {
                    tableFields[i][j] = resultSet.getString(j + 1);
                }
                tableModel.addRow(tableFields[i]);
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        setColumnsWidth(table);
    }

    void initPopupMenu()
    {
        final JPopupMenu popupMenu = new JPopupMenu();
        table.addMouseListener(new MouseAdapter()
        {
            @Override
            public void mouseReleased(MouseEvent e)
            {
                if(e.getButton() == MouseEvent.BUTTON3)
                {
                    popupMenu.show(e.getComponent(), e.getX(), e.getY());
                }
            }
        });
        JMenuItem deleteItem = new JMenuItem("Видалити");
        popupMenu.add(deleteItem);
        deleteItem.addActionListener(e ->
        {
            boolean b = false;
            String id = table.getValueAt(table.getSelectedRow(), 0).toString();
            Double arrival = null;
            Double general = null;

            String purchasePrice = null;
            String supplier = null;
            Date lastDate = null;
            try
            {
                lastDate = main.getDataBaseDateFormat().parse("0000-00-00");
            } catch (ParseException e1) {
                e1.printStackTrace();
            }

            try
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT * FROM `відпускні накладні` " +
                        "WHERE `Номер`='" + noField.getText().trim() + "' " +
                        "AND `Дата`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "';");
                if (resultSet.next())
                {
                    b = true;
                }
                dbi.close(main.k);
            } catch (SQLException e1) {e1.printStackTrace();}

            if (!b)
            {
                try
                {
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("DELETE FROM `прихід` " +
                            "WHERE `Номер`='" + noField.getText().trim() + "' " +
                            "AND `Дата`='" + main.getDataBaseDateFormat().format(datePicker.getDate()) + "'" +
                            "AND `Артикль`='" + id + "';");
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}

                try
                {
                    dbi = new DBI("databassesabc");
                    resultSet = dbi.getSt().executeQuery("SELECT `Прихід`, `Товару в загальному`, `Ціна закупочна`  FROM `товари` " +
                            "WHERE `Артикль`='" + id + "';");
                    if (resultSet.next())
                    {
                        arrival = resultSet.getDouble(1);
                        general = resultSet.getDouble(2);
                        purchasePrice = resultSet.getString(3);
                    }
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}

                try
                {
                    dbi = new DBI("databassesabc");
                    resultSet = dbi.getSt().executeQuery("SELECT `Дата`, `Постачальник`, `Ціна закупочна` FROM `прихід` " +
                            "WHERE `Артикль`='" + id + "' " +
                            "ORDER BY `Дата` DESC " +
                            "LIMIT 1;");
                    if (resultSet.next())
                    {
                        lastDate = resultSet.getDate(1);
                        supplier = resultSet.getString(2);
                        purchasePrice = resultSet.getString(3);
                    }
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}

                try
                {
                    double d1 = arrival - Double.parseDouble(table.getValueAt(table.getSelectedRow(), 4).toString());
                    double d2 = general - Double.parseDouble(table.getValueAt(table.getSelectedRow(), 4).toString());
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("UPDATE `товари` " +
                            "SET " +
                            "`Прихід`='" + (d1) + "', " +
                            "`Товару в загальному`='" + (d2) + "', " +
                            "`Ціна закупочна`='" + purchasePrice + "', " +
                            "`Дата останнього приходу`='" + main.getDataBaseDateFormat().format(lastDate) + "', " +
                            "`Останній постачальник`='" + supplier + "' " +
                            "WHERE `Артикль`='" + id + "';");
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}

                tableModel.removeRow(table.getSelectedRow());
            }
            else
                JOptionPane.showMessageDialog(getThisPanel(), "Видалення цього товару неможливе!", "Помилка", JOptionPane.ERROR_MESSAGE);
        });
    }

    public Invoice getThisPanel()
    {
        return this;
    }

    public JTable getTable()
    {
        return table;
    }

    public JTextField getNoField()
    {
        return noField;
    }

    public DatePicker getDatePicker()
    {
        return datePicker;
    }

    public JScrollPane getScrollPane()
    {
        return scrollPane;
    }

    private class SelectSupplierDialog extends RDialog
    {
        private final JTextField textField;

        public SelectSupplierDialog()
        {
            super(main, "Виберіть постачальника", true);

            GridBagLayout layout = new GridBagLayout();
            setLayout(layout);

            textField = new JTextField(10);
            Collections.sort(supplierArray);
            final JList<String> list = new JList<>(supplierArray.toArray(new String[supplierArray.size()]));
            JButton okBtn = new JButton("Вибрати");
            JButton cancelBtn = new JButton("Закрити");
            textField.getDocument().addDocumentListener(new DocumentListener()
            {
                @Override
                public void insertUpdate(DocumentEvent e)
                {
                    for (int i = 0; i < list.getModel().getSize(); i++)
                    {
                        if (list.getModel().getElementAt(i).contains(textField.getText()))
                        {
                            list.setSelectedValue(list.getModel().getElementAt(i), true);
                            return;
                        }
                    }
                }

                @Override
                public void removeUpdate(DocumentEvent e)
                {
                    for (int i = 0; i < list.getModel().getSize(); i++)
                    {
                        if (list.getModel().getElementAt(i).contains(textField.getText()))
                        {
                            list.setSelectedValue(list.getModel().getElementAt(i), true);
                            return;
                        }
                    }
                }

                @Override
                public void changedUpdate(DocumentEvent e) {}
            });
            okBtn.addActionListener(e ->
            {
                supplier.setText(list.getSelectedValue());
                dispose();
            });
            cancelBtn.addActionListener(e -> dispose());

            add(textField, new GBC(0, 0).setAnchor(GridBagConstraints.WEST));
            add(new JScrollPane(list), new GBC(0, 1).setAnchor(GridBagConstraints.NORTHWEST));
            JPanel btnPanel = new JPanel(new GridBagLayout());
            btnPanel.add(okBtn, new GBC(0, 0));
            btnPanel.add(cancelBtn, new GBC(0, 0));
            add(btnPanel, new GBC(0, 2).setAnchor(GridBagConstraints.EAST));

            pack();
            setResizable(false);
            setLocationRelativeTo(main);
            setLocation(main.getX() + main.getWidth()/2 - this.getWidth()/2, main.getY() + main.getHeight()/2 - this.getHeight()/2);
            setVisible(true);
        }
    }
}
