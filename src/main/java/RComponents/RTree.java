package RComponents;

import Main.Main;
import Support.DBI;
import com.mysql.jdbc.exceptions.MySQLIntegrityConstraintViolationException;
import com.mysql.jdbc.exceptions.MySQLSyntaxErrorException;

import javax.swing.*;
import javax.swing.tree.DefaultMutableTreeNode;
import javax.swing.tree.TreePath;
import java.awt.*;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

public class RTree extends RPanel
{
    private ArrayList<String> warehouse = new ArrayList<>();
    private ArrayList<DefaultMutableTreeNode> warehouseNode = new ArrayList<>();
    private Main main;
    public JTree jt;
    private JPopupMenu popup = new JPopupMenu();

    public RTree(Main main)
    {
        this.main = main;

        setLayout(new BorderLayout());

        popup = new TreePopup();

        setBackground(Color.WHITE);
        /*Properties pr = new Properties();
        File propertyFile = new File("db_conection.propertiess");
        try
        {
            pr.load(new FileReader(propertyFile));
        } catch (IOException e)
        {
            e.printStackTrace();
        }*/
        init();
    }

    public void updatePanel()
    {
        clearPanel();
        init();
        revalidate();
        repaint();

        jt.expandRow(0);
    }

    private void clearPanel()
    {
        this.remove(jt);
    }

    int i;
    private void init()
    {
        DBI dataBase = new DBI("category");
        warehouse.clear();
        warehouseNode.clear();

        DefaultMutableTreeNode level0 = new DefaultMutableTreeNode("Склад");
        jt = new JTree(level0);
        try
        {
            ResultSet dataBaseResult = dataBase.getSt().executeQuery("SHOW TABLES FROM `category`");
            i = -1;
            while (dataBaseResult.next())
            {
                String thisWarehouse = dataBaseResult.getString(1);
                i++;
                warehouse.add(thisWarehouse);
                warehouseNode.add(new DefaultMutableTreeNode(thisWarehouse));
                level0.add(warehouseNode.get(i));
                try
                {
                    DBI warehouseDB = new DBI("category");
                    ResultSet warehouseRS = warehouseDB.getSt().executeQuery("SELECT * FROM `" + smartTrim(thisWarehouse) + "` ORDER BY `Група` ASC");
                    while (warehouseRS.next())
                    {
                        warehouseNode.get(i).add(new DefaultMutableTreeNode(warehouseRS.getString(1)));
                    }
                    warehouseDB.close(main.k);
                    warehouseRS.close();
                } catch (MySQLSyntaxErrorException e1)
                {
                    e1.printStackTrace();
                }
            }
        } catch (SQLException e2)
        {
            e2.printStackTrace();
        }
        jt.addMouseListener(new ClickListener());
        this.add(jt, BorderLayout.WEST);
        dataBase.close(main.k);
    }

    protected RTree getThisTree()
    {
        return this;
    }

    private class ClickListener implements MouseListener
    {
        @Override
        public void mouseClicked(MouseEvent e)
        {
            if(e.getButton() == MouseEvent.BUTTON3)
            {
                popup.show(e.getComponent(), e.getX(), e.getY());
            }
        }

        @Override
        public void mousePressed(MouseEvent e)
        {
            if(e.getButton() == MouseEvent.BUTTON1)
            {
                TreePath treePath = getThisTree().jt.getSelectionPath();
                if(!(treePath == null))
                {
                    if (treePath.getPath().length == 1)
                    {
                        main.getTable().init(null, null);
                    }
                    if (treePath.getPath().length == 2)
                    {
                        main.getTable().init(treePath.getPath()[1].toString(), null);
                    }
                    if (treePath.getPath().length == 3)
                    {
                        main.getTable().init(treePath.getPath()[1].toString(), treePath.getPath()[2].toString());
                    }
                }
            }
        }

        @Override
        public void mouseReleased(MouseEvent e) {}

        @Override
        public void mouseEntered(MouseEvent e) {}

        @Override
        public void mouseExited(MouseEvent e) {}
    }

    private class TreePopup extends JPopupMenu
    {
        TreePath treePath;

        public TreePopup()
        {
            JMenuItem addMenuItem = new JMenuItem("Додати");
            addMenuItem.addActionListener(e ->
            {
                treePath = jt.getSelectionPath();
                String name;
                if(!(treePath == null))
                {
                    if(treePath.getPath().length == 1)
                    {
                        name = JOptionPane.showInputDialog(main, "Будь ласка, введіть назву нового розділу:");
                        if(name != null)
                        {
                            createCategory(name);
                            getThisTree().updatePanel();
                        }

                    }
                    if(treePath.getPath().length == 2)
                    {
                        name = JOptionPane.showInputDialog(main, "Будь ласка, введіть назву нової группи трварів:");
                        createGroup(treePath.getPath()[1].toString(), name);
                        getThisTree().updatePanel();
                    }
                }
            });
            JMenuItem removeMenuItem = new JMenuItem("Видалити");
            removeMenuItem.addActionListener(e ->
            {
                treePath = jt.getSelectionPath();
                if(treePath.getPath().length == 2)
                {
                    String name = treePath.toString().substring(8, treePath.toString().length() - 1);
                    DBI db = new DBI("category");
                    try
                    {
                        db.getSt().execute("DROP TABLE " + name);
                    } catch (SQLException e1)
                    {
                        e1.printStackTrace();
                    }
                    db.close(main.k);
                    getThisTree().updatePanel();
                }
                if(treePath.getPath().length == 3)
                {
                    String name = treePath.getPath()[1].toString();
                    DBI db = new DBI("category");
                    try
                    {
                        db.getSt().execute("DELETE FROM `" + name + "` WHERE `Група`='" + treePath.getPath()[2].toString() + "';");
                    } catch (SQLException e1)
                    {
                        e1.printStackTrace();
                    }
                    db.close(main.k);
                    getThisTree().updatePanel();
                }
            });
            add(addMenuItem);
            add(removeMenuItem);
        }
    }

    public void createGroup(String category, String group)
    {
        createCategory(category);
        DBI db = new DBI("category");
        boolean bool;
        do
        {
            bool = true;
            try
            {
                if(group != null)
                {
                    db.getSt().execute("INSERT INTO `" + category + "` (`Група`) VALUE ('" + group + "')");
                }
            } catch (MySQLIntegrityConstraintViolationException ignored){}
            catch (Exception e1){e1.printStackTrace();
                bool = false;}
            if(!bool)
                JOptionPane.showMessageDialog(main, "Невірна назва групи!");
        } while (!bool);
        db.close(main.k);
        jt.addMouseListener(new ClickListener());
    }

    public void createCategory(String name)
    {
        DBI databasesCon = new DBI("category");
        try
        {
            databasesCon.getSt().execute("CREATE TABLE IF NOT EXISTS `" + name + "` (" +
                    "`Група` text NOT NULL, " +
                    "PRIMARY KEY (`Група`(100)) " +
                    ") ENGINE=InnoDB DEFAULT CHARSET=utf8; ");
        } catch (SQLException e1)
        {
            e1.printStackTrace();
        }
        databasesCon.close(main.k);
        jt.addMouseListener(new ClickListener());
    }
}
