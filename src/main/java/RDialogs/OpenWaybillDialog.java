package RDialogs;

import Main.Main;
import Support.DBI;
import Support.GBC;

import javax.swing.*;
import javax.swing.table.DefaultTableCellRenderer;
import javax.swing.table.TableModel;
import javax.swing.table.TableRowSorter;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.text.ParseException;
import java.util.Comparator;

public class OpenWaybillDialog extends  RDialog
{

    public OpenWaybillDialog(final Main main, String s)
    {
        super(main, "Відкрити накладну", true);

        setLayout(new GridBagLayout());

        DBI dbi;
        ResultSet resultSet;
        if (s.equals("Прихідна накладна"))
        {
            String[] column = new String[]{"Номер", "Дата", "Постачальник"};
            String[][] data = new String[0][];
            try
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " " +
                        "FROM `прихід` " +
                        "GROUP BY `Номер`, `Дата`, `Постачальник`");
                data = new String[getResultSetSize(resultSet) - 1][resultSet.getMetaData().getColumnCount()];
                for (int i = 0; resultSet.next(); i++)
                {
                    for (int j = 0; j < (resultSet.getMetaData().getColumnCount()); j++)
                    {
                        if (j == 1)
                        {
                            data[i][j] = main.getDateFormat().format(main.getDataBaseDateFormat().parse(resultSet.getString(j + 1)));
                        }
                        else
                            data[i][j] = resultSet.getString(j + 1);
                    }
                }
            } catch (SQLException | ParseException e) {e.printStackTrace();}

            final JTable table = new JTable(data, column);
            JButton closeBtn = new JButton("Закрити");
            JButton selectBtn = new JButton("Вибрати");

            TableRowSorter sorter = new TableRowSorter(table.getModel())
            {
                @Override
                public Comparator<?> getComparator(int column)
                {
                    if (column == 0)
                    {
                        return (s1, s2) -> Integer.parseInt((String) s1) - Integer.parseInt((String) s2);
                    }
                    return super.getComparator(column);
                }
            };
            table.setRowSorter(sorter);

            selectBtn.addActionListener(e ->
            {
                try
                {
                    main.getWaybillPanel().getInvoice().initTable(main.getDateFormat().parse(table.getValueAt(table.getSelectedRow(), 1).toString()),
                            table.getValueAt(table.getSelectedRow(), 0).toString(),
                            table.getValueAt(table.getSelectedRow(), 2).toString());
                } catch (ParseException e1) {e1.printStackTrace();}
                dispose();
            });

            closeBtn.addActionListener(e -> dispose());

            add(new JScrollPane(table), new GBC(0, 0, 2, 1));
            add(selectBtn, new GBC(0, 1).setAnchor(GridBagConstraints.WEST));
            add(closeBtn, new GBC(1, 1).setAnchor(GridBagConstraints.EAST));
        }

        if (s.equals("Відпускна накладна"))
        {
            String[] column = new String[]{"Номер відпускної накладної", "Дата відпуску", "Кому відпущена"};
            String[][] data = new String[0][];
            try
            {
                dbi = new DBI("databassesabc");
                resultSet = dbi.getSt().executeQuery("SELECT " + getStringFromArray(column, '`') + " " +
                        "FROM `Відпускні накладні` " +
                        "GROUP BY " + getStringFromArray(column, '`') + ";");
                data = new String[getResultSetSize(resultSet) - 1][resultSet.getMetaData().getColumnCount()];
                for (int i = 0; resultSet.next(); i++)
                {
                    for (int j = 0; j < (resultSet.getMetaData().getColumnCount()); j++)
                    {
                        //if (j == 1)
                        //{
                        //    data[i][j] = main.getDateFormat().format(main.getDataBaseDateFormat().parse(resultSet.getString(j + 1)));
                        //}
                        //else
                        data[i][j] = resultSet.getString(j + 1);
                    }
                }
            } catch (SQLException e) {e.printStackTrace();} //catch (ParseException e) {e.printStackTrace();}

            final JTable table = new JTable(data, column);
            JButton closeBtn = new JButton("Закрити");
            JButton selectBtn = new JButton("Вибрати");

            TableRowSorter sorter = new TableRowSorter(table.getModel())
            {
                @Override
                public Comparator<?> getComparator(int column)
                {
                    if (column == 0)
                    {
                        return (s1, s2) -> Integer.parseInt((String) s1) - Integer.parseInt((String) s2);
                    }
                    return super.getComparator(column);
                }
            };
            table.setRowSorter(sorter);

            selectBtn.addActionListener(e ->
            {
                try {
                    main.getWaybillPanel().getWaybillSelling().initTable(
                            main.getDataBaseDateFormat().parse(table.getValueAt(table.getSelectedRow(), 1).toString()),
                            table.getValueAt(table.getSelectedRow(), 0).toString(),
                            table.getValueAt(table.getSelectedRow(), 2).toString());
                } catch (ParseException e1) {e1.printStackTrace();}
                dispose();
            });

            closeBtn.addActionListener(e -> dispose());

            add(new JScrollPane(table), new GBC(0, 0, 2, 1));
            add(selectBtn, new GBC(0, 1).setAnchor(GridBagConstraints.WEST));
            add(closeBtn, new GBC(1, 1).setAnchor(GridBagConstraints.EAST));
        }

        pack();
        setLocationRelativeTo(main);
        setLocation(main.getX() + main.getWidth()/2 - this.getWidth()/2, main.getY() + main.getHeight()/2 - this.getHeight()/2);
        setVisible(true);
    }
}
