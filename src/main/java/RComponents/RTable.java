package RComponents;

import Main.Main;
import Support.DBI;

import javax.swing.*;
import javax.swing.table.TableColumn;
import javax.swing.table.TableColumnModel;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.awt.event.MouseEvent;
import java.awt.event.MouseListener;
import java.sql.ResultSet;
import java.sql.SQLException;
import java.util.ArrayList;

public class RTable extends RPanel
{
    private JTable table;
    private DBI visibleDB;
    private DBI visibleDBColumn;
    private ResultSet columnNameResult;
    private String[] columnsArray;
    private String[][] tableArray;
    private TablePopup popup = new TablePopup();
    private int selectionColumn;
    private ResultSet columnRS;
    private ArrayList<String> invisibleColumnArray = new ArrayList<>();
    private ArrayList<String> visibleColumnArray = new ArrayList<>();
    private Main main = null;

    public String getCategory() {
        return category;
    }

    public String getGroup() {

        return group;
    }

    private String category;
    private String group;

    public RTable(Main main)
    {
        this.main = main;
        setLayout(new BorderLayout());
        setBackground(Color.WHITE);
    }

    public void init(String category, String group)
    {
        this.removeAll();

        this.category = category;
        this.group = group;

        DBI db = new DBI("databassesabc");
        visibleDB = new DBI("databassesabc");
        visibleDBColumn = new DBI("databassesabc");

        visibleColumnArray.clear();
        invisibleColumnArray.clear();

        try
        {
            visibleDBColumn = new DBI("databassesabc");
            columnNameResult =  visibleDBColumn.getSt().executeQuery("SHOW COLUMNS FROM `columnvisible`");
            ResultSet columnVisibleRS = visibleDB.getSt().executeQuery("SELECT * FROM `columnvisible`");
            columnVisibleRS.next();
            for(int i = 1; columnNameResult.next(); i++)
            {
                if(columnVisibleRS.getString(i).equals("true"))
                {
                    visibleColumnArray.add(columnNameResult.getString(1));
                }
                if(columnVisibleRS.getString(i).equals("false"))
                {
                    invisibleColumnArray.add(columnNameResult.getString(1));
                }
            }
            columnsArray  = convertArray(visibleColumnArray);
        } catch (SQLException e)
        {
            e.printStackTrace();
        }

        try
        {
            ResultSet tableResult;
            if (group == null)
            {
                if (category == null)
                {
                    tableResult = db.getSt().executeQuery("SELECT " + getStringFromArray(visibleColumnArray, '`') + "FROM `товари`;");
                }
                else
                {
                    tableResult = db.getSt().executeQuery("SELECT " + getStringFromArray(visibleColumnArray, '`') + "FROM `товари` " +
                            "WHERE `Категорія`='" + smartTrim(category) + "';");
                }
            }
            else
            {
                tableResult = db.getSt().executeQuery("SELECT " + getStringFromArray(visibleColumnArray, '`') + "FROM `товари` " +
                        "WHERE `Категорія`='" + smartTrim(category) + "' AND `Група`='" + smartTrim(group) + "';");
            }
            tableArray = new String[getResultSetSize(tableResult) - 1][tableResult.getMetaData().getColumnCount()];
            for(int i = 0; tableResult.next(); i++)
            {
                for(int j = 0; j< tableResult.getMetaData().getColumnCount(); j++)
                {
                    if (!columnsArray[j].equals("Дата останнього приходу"))
                    {
                        tableArray[i][j] = tableResult.getString(j + 1);
                    }
                    else
                    try
                    {
                        tableArray[i][j] = main.getDateFormat().format(tableResult.getDate(j + 1));
                    } catch(SQLException e1)
                    {
                        tableArray[i][j] = "00.00.0000";
                    }
                }
            }
        } catch (SQLException e) {e.printStackTrace();}

        table = new JTable(tableArray, columnsArray);

        this.add(new JScrollPane(table), BorderLayout.CENTER);
        table.getTableHeader().addMouseListener(new ColumnListener());
        this.revalidate();

        db.close(main.k);
        visibleDB.close(main.k);
        visibleDBColumn.close(main.k);

        setColumnsWidth(table);
    }

    public void removeTableColumn(String column)
    {
        TableColumnModel tcm = table.getColumnModel();
        int i = tcm.getColumnIndex(column);
        TableColumn tc = tcm.getColumn(i);
        table.getColumnModel().removeColumn(tc);
    }

    public void removeTableColumn(int columnIndex)
    {
        TableColumnModel tcm = table.getColumnModel();
        int i = columnIndex;
        TableColumn tc = tcm.getColumn(i);
        table.getColumnModel().removeColumn(tc);
    }

    private class ColumnListener implements MouseListener
    {
        @Override
        public void mouseClicked(MouseEvent e)
        {
            if(e.getButton() == MouseEvent.BUTTON3)
            {
                selectionColumn = table.getTableHeader().columnAtPoint(e.getPoint());
                popup.show(table.getTableHeader().getComponentAt(e.getX(), e.getY()), e.getX(), getY());
            }
        }

        @Override
        public void mousePressed(MouseEvent e) {
        }

        @Override
        public void mouseReleased(MouseEvent e) {
        }

        @Override
        public void mouseEntered(MouseEvent e) {
        }

        @Override
        public void mouseExited(MouseEvent e) {
        }
    }

    private class TablePopup extends JPopupMenu
    {
        public TablePopup()
        {
            JMenuItem addColumn = new JMenuItem("Додати колонку");
            JMenuItem removeColumn = new JMenuItem("Видалити колонку");
            addColumn.addActionListener(e ->
            {
                try
                {
                    visibleDBColumn = new DBI("databassesabc");
                    visibleDB = new DBI("databassesabc");
                    columnNameResult =  visibleDBColumn.getSt().executeQuery("SHOW COLUMNS FROM `columnvisible`");
                    columnRS = visibleDB.getSt().executeQuery("SELECT * FROM `columnvisible`");
                    columnRS.next();
                    invisibleColumnArray.clear();
                    for(int i = 1; columnNameResult.next(); i++)
                    {
                        if(columnRS.getString(i).equals("false"))
                        {
                            invisibleColumnArray.add(columnNameResult.getString(1));
                        }
                    }
                    String name = null;
                    if(invisibleColumnArray.size() != 0)
                    {
                        name = (String) JOptionPane.showInputDialog(main,
                                "Виберіть, яку колонку додати:",
                                "Додовання колонки",
                                JOptionPane.PLAIN_MESSAGE,
                                UIManager.getIcon("OptionPane.q"),
                                convertArray(invisibleColumnArray),
                                invisibleColumnArray.get(0));
                    }
                    if(!(name == null))
                    {
                        visibleDB.getSt().execute("UPDATE columnvisible SET `" + name + "` = \"true\"");
                        init(getCategory(), getGroup());
                    }
                } catch (SQLException e1)
                {
                    e1.printStackTrace();
                }
            });
            removeColumn.addActionListener(e ->
            {
                try
                {
                    visibleDB = new DBI("databassesabc");
                    visibleDB.getSt().execute("UPDATE columnvisible " +
                            "SET `" + table.getColumnModel().getColumn(selectionColumn).getIdentifier().toString() + "` = 'false'");
                } catch (SQLException e1)
                {
                    e1.printStackTrace();
                }
                removeTableColumn(selectionColumn);
            });
            add(addColumn);
            add(removeColumn);
        }
    }

    public JTable getTable()
    {
        return table;
    }
}
