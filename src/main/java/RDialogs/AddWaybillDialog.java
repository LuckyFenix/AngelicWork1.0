package RDialogs;

import Main.Main;
import Support.DBI;
import Support.GBC;

import javax.swing.*;
import javax.swing.table.DefaultTableModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;
import java.util.Date;

public class AddWaybillDialog extends RDialog
{
    private final String no;
    private final Date date;
    private final DefaultTableModel leftTableModel;
    private final DefaultTableModel rightTableModel;
    private final String[] column;
    private final JTable leftTable;
    private final JTable rightTable;
    private DBI dbi;
    private ArrayList<String> row;
    private final Main main;
    private final String staff;

    public AddWaybillDialog(String no, Date date, String staff, Main main)
    {
        super(main, "Додати накладні", true);
        this.main = main;

        GridBagLayout layout = new GridBagLayout();
        setLayout(layout);

        this.no = no;
        this.date = date;
        this.staff = staff;

        column = new String[]{"Номер", "Дата", "Постачальник"};
        leftTableModel = new DefaultTableModel(new String[0][0], column);
        rightTableModel = new DefaultTableModel(new String[0][0], column);

        leftTable = new JTable(leftTableModel);
        rightTable = new JTable(rightTableModel);
        JButton toLeftBtn = new JButton("<<");
        JButton toRightBtn = new JButton(">>");
        JButton okBtn = new JButton("Підтвердити");
        JButton cancelBtn = new JButton("Закрити");

        initLeftTable();

        toRightBtn.addActionListener(e ->
        {
            for (int j = 0; j < column.length; j++)
            {
                row.add(leftTable.getValueAt(leftTable.getSelectedRow(), j).toString());
            }
            rightTableModel.addRow(row.toArray());
            row.clear();
            leftTableModel.removeRow(leftTable.getSelectedRow());
            setColumnsWidth(rightTable);
        });

        toLeftBtn.addActionListener(e ->
        {
            for (int j = 0; j < column.length; j++)
            {
                row.add(rightTable.getValueAt(rightTable.getSelectedRow(), j).toString());
            }
            leftTableModel.addRow(row.toArray());
            rightTableModel.removeRow(rightTable.getSelectedRow());
            setColumnsWidth(rightTable);
        });

        okBtn.addActionListener(new OkBtnListener());
        cancelBtn.addActionListener(e -> dispose());

        JPanel panel = new JPanel(new GridBagLayout());
        panel.add(new JScrollPane(leftTable), new GBC(0, 0, 1, 3).setAnchor(GridBagConstraints.NORTHWEST));
        panel.add(toRightBtn, new GBC(1, 0).setAnchor(GridBagConstraints.NORTH));
        panel.add(toLeftBtn, new GBC(1, 1).setAnchor(GridBagConstraints.NORTH));
        panel.add(new JScrollPane(rightTable), new GBC(2, 0, 1, 3).setAnchor(GridBagConstraints.NORTHWEST));

        JPanel btnPanel = new JPanel(new GridBagLayout());
        btnPanel.add(okBtn, new GBC(0, 0));
        btnPanel.add(cancelBtn, new GBC(1, 0));

        panel.add(btnPanel, new GBC(0, 3, 3, 1).setAnchor(GridBagConstraints.EAST));

        add(panel, new GBC(0, 0).setAnchor(GridBagConstraints.NORTHWEST).setFill(20));

        pack();

        setResizable(false);
        setLocationRelativeTo(main);
        setLocation(main.getX() + main.getWidth() / 2 - this.getWidth() / 2, main.getY() + main.getHeight() / 2 - this.getHeight() / 2);
        revalidate();
        setVisible(true);
    }

    private void initLeftTable()
    {
        row = new ArrayList<>();

        try
        {
            dbi = new DBI("databassesabc");
            ResultSet resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " FROM `прихід` " +
                    "WHERE `Була відпущена`='0' " +
                    "GROUP BY `Номер`, `Дата`, `Постачальник`;");
            while (resultSet.next())
            {
                for (int i = 1; i <= column.length; i++)
                    row.add(resultSet.getString(i));
                leftTableModel.addRow(row.toArray());
                row.clear();
            }
        } catch (SQLException e) {e.printStackTrace();}
        setColumnsWidth(leftTable);
    }

    private class OkBtnListener implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            for (int i = 0; i < rightTable.getRowCount(); i++)
            {
                try
                {
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("UPDATE `прихід` SET `Була відпущена`='1' " +
                            "WHERE `Номер`='" + rightTable.getValueAt(i, 0) + "' " +
                            "AND `Дата`='" + rightTable.getValueAt(i, 1) + "';");
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}

                try
                {
                    dbi = new DBI("databassesabc");
                    dbi.getSt().execute("INSERT INTO `відпускні накладні` " +
                            "(`Дата відпуску`, `Номер відпускної накладної`, `Дата`, `Номер`, `Постачальник`, `Кому відпущена`) VALUES " +
                            "('" + main.getDataBaseDateFormat().format(date) + "', " +
                            "'" + no + "', " +
                            "'" + rightTable.getValueAt(i, 1) + "', " +
                            "'" + rightTable.getValueAt(i, 0) + "', " +
                            "'" + rightTable.getValueAt(i, 2) + "', " +
                            "'" + staff + "');");
                    dbi.close(main.k);
                } catch (SQLException e1) {e1.printStackTrace();}
            }
            main.getWaybillPanel().getWaybillSelling().initTable();
            dispose();
        }
    }
}
