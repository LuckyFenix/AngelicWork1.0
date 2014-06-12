package RDialogs;

import Main.Main;
import Support.DBI;
import Support.GBC;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.sql.ResultSet;
import java.sql.SQLException;

public class StaffDialog extends RDialog
{
    private Main main;
    private final JList<String> list;
    private DBI staffDB;
    private final DefaultListModel<String> listModel = new DefaultListModel<String>();
    private final JScrollPane scrollPane;

    public StaffDialog(Main main)
    {
        super(main, "Працівники", true);
        this.main = main;

        GridBagLayout layout = new GridBagLayout();
        setLayout(layout);

        ResultSet resultSet;
        staffDB = new DBI("databassesabc");
        try
        {
            resultSet = staffDB.getSt().executeQuery("SELECT * FROM `працівники`");
            String s;
            while (resultSet.next())
            {
                s = resultSet.getString(1);
                listModel.addElement(s);
            }
        } catch (SQLException e) {e.printStackTrace();}

        list = new JList<>(listModel);
        JButton newStaffBtn = new JButton("Новий");
        JButton deleteStaffBtn = new JButton("Видалити");
        JButton editStaffBtn = new JButton("Редактувати");

        newStaffBtn.addActionListener(new newStaffAction());
        deleteStaffBtn.addActionListener(new deleteStaffAction());
        editStaffBtn.addActionListener(new editStaffAction());
        scrollPane = new JScrollPane(list);

        add(scrollPane, new GBC(0, 0, 3, 1).setAnchor(GBC.WEST).setIpad(125, 100));
        add(newStaffBtn, new GBC(0, 1).setAnchor(GBC.WEST));
        add(deleteStaffBtn, new GBC(1, 1).setAnchor(GBC.CENTER));
        add(editStaffBtn, new GBC(2, 1).setAnchor(GBC.EAST));

        pack();

        setResizable(false);
        setLocationRelativeTo(main);
        setLocation(main.getX() + main.getWidth()/2 - this.getWidth()/2, main.getY() + main.getHeight()/2 - this.getHeight()/2);
        revalidate();
        setVisible(true);
        staffDB.close(main.k);
    }

    private class newStaffAction implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            staffDB = new DBI("databassesabc");
            String name = JOptionPane.showInputDialog(getThisDialog(), "Введіть ім'я співробітника:");
            try
            {
                staffDB.getSt().execute("INSERT INTO `працівники` (`Імя`) VALUE ('" + name + "');");
            } catch (SQLException e1) {e1.printStackTrace();}
            listModel.addElement(name);
            int index = listModel.size() - 1;
            list.setSelectedIndex(index);
            list.ensureIndexIsVisible(index);
            staffDB.close(main.k);
            scrollPane.revalidate();
        }
    }

    private class deleteStaffAction implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            if(list.getSelectedIndex() != -1)
            {
                staffDB = new DBI("databassesabc");
                try
                {
                    staffDB.getSt().execute("DELETE FROM `працівники` WHERE `Імя`='" + listModel.getElementAt(list.getSelectedIndex()) + "';");
                } catch (SQLException e1) {e1.printStackTrace();}
                listModel.remove(list.getSelectedIndex());
                staffDB.close(main.k);
            }
        }
    }

    private class editStaffAction implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            staffDB = new DBI("databassesabc");
            int i = list.getSelectedIndex();
            String start = listModel.getElementAt(i);
            String end = JOptionPane.showInputDialog(main, "", start);
            try {
                staffDB.getSt().execute("UPDATE `працівники` SET `Імя`='" + end + "' WHERE `Імя`='" + start + "';");
            } catch (SQLException e1) {
                e1.printStackTrace();
            }
            listModel.setElementAt(end, i);
        }
    }

    public StaffDialog getThisDialog()
    {
        return this;
    }
}
