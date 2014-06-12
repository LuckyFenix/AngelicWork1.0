package RDialogs;

import Main.Main;
import RComponents.IndividArrayList;
import RComponents.RComboBox;
import Support.DBI;
import Support.GBC;

import javax.swing.*;
import javax.swing.event.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.*;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Collections;
import java.util.Date;

/*
Диаглог добавления новой номенклатуры в накладную прихода
 */

public class InvoiceNewNomenclatureDialog extends RDialog
{
    private final Main main;
    private RComboBox categoryBox;
    private RComboBox groupBox;
    private RComboBox nameBox;
    private DBI dbi;
    private ResultSet resultSet;
    private ArrayList<String> categoryList = new ArrayList<>();
    private JTextField idField = new JTextField(13);
    private final JTextField unitOfMeasureField = new JTextField(3);
    private final JButton newBtn;
    private final JTextField purchasePriceField = new JTextField(10);
    private final JTextField amountField = new JTextField(10);
    private final JTextField sellingPriceField = new JTextField(10);
    private final JTextField sumField = new JTextField(10);
    private final JTextField reservationField = new JTextField(6);
    private String newName = null;
    private int k1 = 0;
    private int k2 = 0;
    private int k3 = 0;
    private final String no;
    private final String supplier;
    private final Date date;
    private DefaultTableModel tableModel;
    private JTable table;
    private final String[] column;


    public InvoiceNewNomenclatureDialog(Main main, String no, String supplier, Date date)
    {
        super(main, "Нова номенклатура", true);
        this.main = main;
        this.no = no;
        this.supplier = supplier;
        this.date = date;

        GridBagLayout layout = new GridBagLayout();
        setLayout(layout);

        JPanel panel = new JPanel(new GridBagLayout());
        JPanel btnPanel = new JPanel(new GridBagLayout());
        JPanel namePanel = new JPanel(new GridBagLayout());
        categoryBox = new RComboBox(initCategoryBox());
        groupBox = new RComboBox(initGroupBox(categoryBox.getSelectedItem().toString()));
        nameBox = new RComboBox(initNameBox(categoryBox.getSelectedItem().toString(), groupBox.getSelectedItem().toString()));
        idField.setText("" + initIdField(nameBox.getSelectedItem().toString()));
        newBtn = new JButton("Нова");
        JButton saveBtn = new JButton("Додати");
        JButton cancelBtn = new JButton("Закрити");

        column = new String[]{"Дата", "Постачальник", "Ціна закупочна", "Кількість", "Ціна", "Сума"};
        tableModel = new DefaultTableModel(new String[][]{}, column);
        table = new JTable(tableModel);

        sumField.setForeground(Color.RED);
        initValue();
        initTable();

        categoryBox.addItemListener(new CategoryBoxListener());
        groupBox.addItemListener(new GroupBoxListener());
        nameBox.addItemListener(new NameBoxListener());
        idField.addKeyListener(new IdFieldListener());
        newBtn.addActionListener(new NewBtnListener());
        amountField.getDocument().addDocumentListener(new AmountFieldListener());
        saveBtn.addActionListener(new SaveBtnListener());
        cancelBtn.addActionListener(e -> getThisDialog().dispose());

        panel.add(new JLabel("Категорія:"), new GBC(0, 0).setAnchor(GBC.WEST).setInsets(0, 0, 0, 10));
        panel.add(categoryBox, new GBC(1, 0, 2, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Група:"), new GBC(0, 1).setAnchor(GBC.WEST));
        panel.add(groupBox, new GBC(1, 1, 2, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Артикль:"), new GBC(0, 2).setAnchor(GBC.WEST));
        panel.add(idField, new GBC(1, 2, 2, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Назва товару"), new GBC(0, 3).setAnchor(GBC.WEST));

        namePanel.add(nameBox, new GBC(1, 0).setAnchor(GBC.WEST));
        namePanel.add(newBtn, new GBC(2, 0));
        panel.add(namePanel, new GBC(1, 3, 2, 1));

        panel.add(new JLabel("Од. вимірювання:"), new GBC(0, 4).setAnchor(GBC.WEST));
        panel.add(unitOfMeasureField, new GBC(1, 4, 1, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Ціна закупочна:"), new GBC(0, 5).setAnchor(GBC.WEST));
        panel.add(purchasePriceField, new GBC(1, 5, 1, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Кількість:"), new GBC(0, 6).setAnchor(GBC.WEST));
        panel.add(amountField, new GBC(1, 6, 1, 1).setAnchor(GBC.EAST));
        panel.add(new JLabel("Ціна продажна:"), new GBC(0, 7).setAnchor(GBC.NORTHWEST));
        panel.add(sellingPriceField, new GBC(1, 7, 1, 1).setAnchor(GBC.NORTHEAST));
        panel.add(new JLabel("Сума:"), new GBC(0, 8).setAnchor(GBC.NORTHWEST));
        panel.add(sumField, new GBC(1, 8, 1, 1).setAnchor(GBC.NORTHEAST));
        panel.add(new JLabel("Бронь:"), new GBC(2, 5).setAnchor(GBC.CENTER));
        panel.add(reservationField, new GBC(2, 6, 1, 1).setAnchor(GBC.CENTER));

        btnPanel.add(saveBtn, new GBC(0, 0).setAnchor(GBC.CENTER));
        btnPanel.add(cancelBtn, new GBC(1, 0).setAnchor(GBC.CENTER).setInsets(0, 5, 0, 0));

        panel.add(btnPanel, new GBC(0, 9, 4, 1).setAnchor(GridBagConstraints.EAST));

        add(panel, new GBC(0, 0).setAnchor(GBC.CENTER).setInsets(20));
        add(new JScrollPane(table), new GBC(1, 0));

        pack();
        setResizable(false);
        setLocationRelativeTo(main);
        setLocation(main.getX() + main.getWidth()/2 - this.getWidth()/2, main.getY() + main.getHeight()/2 - this.getHeight()/2);
        setVisible(true);
    }

    private void initTable()
    {
        for (int i = 0; i < table.getRowCount(); i++)
        {
            tableModel.removeRow(0);
        }
        table.removeAll();
        try
        {
            ArrayList<String> row = new ArrayList<String>();
            dbi = new DBI("databassesabc");
            resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " " +
                    "FROM `прихід` " +
                    "WHERE `Назва товару`='" + nameBox.getSelectedItem().toString() + "' " +
                    "ORDER BY `Дата`;");
            while (resultSet.next())
            {
                for (int i = 1; i <= resultSet.getMetaData().getColumnCount(); i++)
                    row.add(resultSet.getString(i));
                tableModel.addRow(row.toArray());
                row.clear();
            }
        } catch (SQLException e) {e.printStackTrace();}
        setColumnsWidth(table);
        revalidate();
    }

    protected InvoiceNewNomenclatureDialog getThisDialog()
    {
        return this;
    }

    private void initValue()
    {
        String name = nameBox.getSelectedItem().toString();
        String unitOfMeas = null;
        String purchasePrice = null;
        String sellingPrice = null;
        String reservationValue = null;
        double sum = 0;
        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery("SELECT `Одиниці вимірювання`, `Ціна закупочна`, `Ціна`, `Бронь` " +
                    "FROM `товари` " +
                    "WHERE `Назва товару`='" + name + "' " +
                    "ORDER BY `Дата останнього приходу` DESC;");
            if(resultSet.next())
            {
                unitOfMeas = resultSet.getString(1);
                purchasePrice = resultSet.getString(2);
                sellingPrice = resultSet.getString(3);
                reservationValue = resultSet.getString(4);
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        amountField.setText("");
        unitOfMeasureField.setText(unitOfMeas);
        purchasePriceField.setText(purchasePrice);
        sellingPriceField.setText(sellingPrice);
        sumField.setText("" + sum);
        reservationField.setText(reservationValue);
    }

    private int initIdField(String name)
    {
        int id = 0;
        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery("SELECT `Артикль` FROM `товари` WHERE `Назва товару`='" + name + "';");
            if(resultSet.next())
                id = Integer.parseInt(resultSet.getString(1));
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        return id;
    }

    private ArrayList<String> initNameBox(String category, String group)
    {
        IndividArrayList<String> arrayList = new IndividArrayList<>();
        String query;
        if(group == null)
        {
            if(category == null)
            {
                query = "SELECT `Назва товару` FROM `товари`;";
            }
            else
                query = "SELECT `Назва товару` FROM `товари` WHERE `Категорія`='" + category + "';";
        }
        else
            query = "SELECT `Назва товару` FROM `товари` WHERE `Категорія`='" + category + "' AND `Група`='" + group + "';";

        dbi = new DBI("databassesabc");
        try
        {
            resultSet = dbi.getSt().executeQuery(query);
            while (resultSet.next())
            {
                arrayList.add(resultSet.getString(1));
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        return arrayList;
    }

    private ArrayList<String> initGroupBox(String category)
    {
        ArrayList<String> arrayList = new ArrayList<>();
        String query;
        query = "SELECT * FROM `" + category + "`;";
        dbi = new DBI("category");
        try
        {
            resultSet = dbi.getSt().executeQuery(query);
            while (resultSet.next())
            {
               arrayList.add(resultSet.getString(1));
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}

        Collections.sort(arrayList);
        return arrayList;
    }

    private ArrayList<String> initCategoryBox()
    {
        final ArrayList<String> arrayList = new ArrayList<>();
        dbi = new DBI("category");
        try
        {
            resultSet = dbi.getSt().executeQuery("SHOW TABLES FROM `category`;");
            while (resultSet.next())
            {
                arrayList.add(resultSet.getString(1));
                categoryList.add(resultSet.getString(1));
            }
            dbi.close(main.k);
        } catch (SQLException e) {e.printStackTrace();}
        return arrayList;
    }

    public class IdFieldListener extends KeyAdapter
    {
        public void keyPressed(KeyEvent e)
        {
            if (e.getKeyCode() == KeyEvent.VK_ENTER)
            {
                dbi = new DBI("databassesabc");
                try
                {
                    resultSet = dbi.getSt().executeQuery("SELECT `Категорія`, `Група`, `Назва товару` FROM `товари` WHERE `Артикль`='" + idField.getText() + "';");
                    resultSet.next();
                    String category = resultSet.getString(1);
                    String group = resultSet.getString(2);
                    String name = resultSet.getString(3);
                    categoryBox.setSelectedItem(Character.toLowerCase(category.charAt(0)) + category.substring(1));
                    //groupBox.changeComboBox(initGroupBox(categoryBox.getSelectedItem().toString()));
                    groupBox.setSelectedItem(group);
                    //nameBox.changeComboBox(initNameBox(categoryBox.getSelectedItem().toString(), groupBox.getSelectedItem().toString()));
                    nameBox.setSelectedItem(name);
                    initValue();
                    dbi.close(main.k);
                } catch (Exception e1)
                {
                    JOptionPane.showMessageDialog(getThisDialog(), "Товару під таким артиклем не існує!", "Помилка", JOptionPane.ERROR_MESSAGE);
                }
                getThisDialog().pack();
            }
        }
    }

    public class NameBoxListener implements ItemListener
    {
        @Override
        public void itemStateChanged(ItemEvent e)
        {
            k1++;
            if(nameBox.getItemCount() != 0 && k1%2 == 0)
            {
                idField.setText("" + initIdField(nameBox.getSelectedItem().toString()));
                initValue();
                initTable();
                getThisDialog().pack();
            }
        }
    }

    public class GroupBoxListener implements ItemListener
    {
        @Override
        public void itemStateChanged(ItemEvent e)
        {
            k2++;
            if(groupBox.getItemCount() != 0 && k2%2 == 0)
            {
                String category = categoryBox.getSelectedItem().toString();
                String group = groupBox.getSelectedItem().toString();
                nameBox.changeComboBox(initNameBox(category, group));
                idField.setText("" + initIdField(nameBox.getSelectedItem().toString()));
                initValue();
                getThisDialog().pack();
            }
        }
    }

    public class CategoryBoxListener implements ItemListener
    {
        @Override
        public void itemStateChanged(ItemEvent e)
        {
            k3++;
            if(k3%2 == 0)
            {
                groupBox.changeComboBox(initGroupBox(categoryBox.getSelectedItem().toString()));
                idField.setText("" + initIdField(nameBox.getSelectedItem().toString()));
                initValue();
                getThisDialog().pack();
            }
        }
    }

    private class NewBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            if (newBtn.getText().equals("Нова"))
            {
                newName = JOptionPane.showInputDialog(getThisDialog(), "Введіть назву нового товару:");
                if (newName != null)
                {
                    ArrayList<String> arrayList = nameBox.getItemArray();
                    arrayList.add(newName);
                    nameBox.changeComboBox(arrayList);
                    nameBox.setSelectedItem(newName);
                    nameBox.setEnabled(false);
                    groupBox.setEnabled(false);
                    categoryBox.setEnabled(false);
                    idField.setEditable(false);
                    newBtn.setText("Відмінити");
                    getThisDialog().pack();
                }
            }
            else
            {
                ArrayList<String> arrayList = nameBox.getItemArray();
                arrayList.remove(newName);
                nameBox.changeComboBox(arrayList);
                nameBox.setEnabled(true);
                groupBox.setEnabled(true);
                categoryBox.setEnabled(true);
                idField.setEditable(true);
                newBtn.setText("Нова");
                initValue();
                getThisDialog().pack();
            }

        }
    }

    private class AmountFieldListener implements DocumentListener
    {
        @Override
        public void insertUpdate(DocumentEvent e)
        {
            try
            {
                sumField.setText("" + (Double.parseDouble(amountField.getText()) * Double.parseDouble(purchasePriceField.getText())));
            } catch (NumberFormatException ignored){}
        }

        @Override
        public void removeUpdate(DocumentEvent e)
        {
            try
            {
                if (!amountField.getText().equals(""))
                    sumField.setText("" + (Double.parseDouble(amountField.getText()) * Double.parseDouble(purchasePriceField.getText())));
                else
                    sumField.setText("" + Double.parseDouble("0"));
            } catch (NumberFormatException ignored){}
        }

        @Override
        public void changedUpdate(DocumentEvent e) {}
    }

    private class SaveBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            try
            {
                String id;
                String arrival;
                String general;

                if (reservationField.getText().trim().equals(""))
                    reservationField.setText("0");
                if (unitOfMeasureField.getText().trim().equals(""))
                    unitOfMeasureField.setText("шт.");
                if (sellingPriceField.getText().trim().equals(""))
                    sellingPriceField.setText("0");
                if (purchasePriceField.getText().trim().equals(""))
                    purchasePriceField.setText("0");

                if(!newBtn.getText().equals("Нова"))
                {
                    String category = categoryBox.getSelectedItem().toString();

                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("INSERT INTO `товари` " +
                            "(`Категорія`, `Група`, `Назва товару`, `Одиниці вимірювання`, `Ціна`, `Ціна закупочна`) VALUES " +
                            "('" + (Character.toUpperCase(category.charAt(0)) + category.substring(1, category.length())) + "', " +
                            "'" + groupBox.getSelectedItem().toString() + "', " +
                            "'" + nameBox.getSelectedItem().toString() + "', " +
                            "'" + unitOfMeasureField.getText().trim() + "', " +
                            "'" + sellingPriceField.getText().trim() + "', " +
                            "'" + purchasePriceField.getText().trim() + "');");
                    dbi.close(main.k);
                    nameBox.setEnabled(true);
                    groupBox.setEnabled(true);
                    categoryBox.setEnabled(true);
                    idField.setEditable(true);
                    newBtn.setText("Нова");
                }

                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT `Артикль`, `Прихід`, `Товару в загальному` " +
                        "FROM `товари` " +
                        "WHERE `Назва товару`='" + smartTrim(nameBox.getSelectedItem().toString()) + "' " +
                        "AND `Ціна`='" + Double.parseDouble(sellingPriceField.getText().trim()) + "';");
                if (resultSet.next())
                {
                    id = resultSet.getString(1);
                    arrival = resultSet.getString(2);
                    general = resultSet.getString(3);
                }
                else
                {
                    DBI db = new DBI("databassesabc");
                    db.getSt().execute("INSERT INTO `товари` " +
                            "(`Категорія`, `Група`, `Назва товару`, `Одиниці вимірювання`, `Ціна`, `Ціна закупочна`) VALUES " +
                            "('" + categoryBox.getSelectedItem().toString() + "', " +
                            "'" + groupBox.getSelectedItem().toString() + "', " +
                            "'" + nameBox.getSelectedItem().toString() + "', " +
                            "'" + unitOfMeasureField.getText().trim() + "', " +
                            "'" + sellingPriceField.getText().trim() + "', " +
                            "'" + purchasePriceField.getText().trim() + "');");
                    db.close(main.k);

                    dbi = new DBI("databassesabc");
                    resultSet = dbi.getSt().executeQuery("SELECT `Артикль` " +
                            "FROM `товари` " +
                            "WHERE `Назва товару`='" + smartTrim(nameBox.getSelectedItem().toString()) + "' " +
                            "AND `Ціна`='" + Double.parseDouble(sellingPriceField.getText().trim()) + "';");
                    resultSet.next();
                    id = resultSet.getString(1);
                    arrival = "0";
                    general = "0";
                }
                dbi.close(main.k);

                dbi = new DBI("databassesabc");
                dbi.getSt().execute("INSERT INTO `прихід` " +
                        "(`Номер`, `Дата`, `Артикль`, `Назва товару`, `Одиниці вимірювання`, `Постачальник`, `Ціна закупочна`, `Кількість`, `Ціна`, `Сума`) VALUES " +
                        "('" + no + "', " +
                        "'" + main.getDataBaseDateFormat().format(date) + "', " +
                        "'" + id + "', " +
                        "'" + nameBox.getSelectedItem().toString() + "', " +
                        "'" + unitOfMeasureField.getText().trim() + "', " +
                        "'" + supplier + "', " +
                        "'" + purchasePriceField.getText().trim() + "', " +
                        "'" + amountField.getText().trim() + "', " +
                        "'" + sellingPriceField.getText().trim() + "', " +
                        "'" + sumField.getText().trim() + "');");
                dbi.close(main.k);

                dbi = new DBI("databassesabc");
                dbi.getSt().execute("UPDATE `товари` SET " +
                        "`Прихід`='" + (Double.parseDouble(arrival) + Double.parseDouble(amountField.getText().trim())) + "', " +
                        "`Товару в загальному`='" + (Double.parseDouble(general) + Double.parseDouble(amountField.getText().trim())) + "', " +
                        "`Ціна закупочна`='" + purchasePriceField.getText().trim() + "', " +
                        "`Бронь`='" + reservationField.getText().trim() + "', " +
                        "`Дата останнього приходу`='" + main.getDataBaseDateFormat().format(date) + "', " +
                        "`Останній постачальник`='" + supplier + "' " +
                        "WHERE `Артикль`='" + id + "';");
                dbi.close(main.k);

                main.getWaybillPanel().getInvoice().initTable();
                main.getWaybillPanel().getInvoice().getTable().setRowSelectionInterval(main.getWaybillPanel().getInvoice().getTable().getRowCount() - 1, main.getWaybillPanel().getInvoice().getTable().getRowCount() - 1);
                main.getWaybillPanel().getInvoice().getScrollPane().getVerticalScrollBar().setValue(main.getWaybillPanel().getInvoice().getScrollPane().getVerticalScrollBar().getMaximum());

                amountField.setText("");
            } catch (Exception e1)
            {
                JOptionPane.showMessageDialog(getThisDialog(), e1.getMessage() + " " + e1.getLocalizedMessage(), "Ошибка", JOptionPane.ERROR_MESSAGE);
                e1.printStackTrace();
            }
        }
    }
}
