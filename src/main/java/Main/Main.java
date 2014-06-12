package Main;

import RComponents.*;

import javax.swing.*;
import java.awt.BorderLayout;
import java.text.SimpleDateFormat;

public class Main extends JFrame
{
    public SimpleDateFormat dateFormat = new SimpleDateFormat("dd.MM.yyyy");
    public SimpleDateFormat dataBaseDateFormat = new SimpleDateFormat("yyyy-MM-dd");
    private WaybillPanel waybillPanel;
    private RTree tree;
    private RTable table;
    private RToolBar toolBat;
    private RMenuBar menuBar;
    public int k = 0;

    public Main()
    {
        UIManager.LookAndFeelInfo[] infos = UIManager.getInstalledLookAndFeels();
        try
        {
            UIManager.setLookAndFeel(infos[1].getClassName());
        } catch (Exception e)
        {
            e.printStackTrace();
        }

        setSize(1366, 768);
        setExtendedState(MAXIMIZED_BOTH);

        toolBat = new RToolBar(this);
        menuBar = new RMenuBar(this);
        waybillPanel = new WaybillPanel(this);
        tree = new RTree(this);
        table = new RTable(this);

        setJMenuBar(menuBar);

        JSplitPane horizontalSP = new JSplitPane(JSplitPane.HORIZONTAL_SPLIT);
        horizontalSP.setLeftComponent(new JScrollPane(tree));
        horizontalSP.setRightComponent(new JScrollPane(table));

        JSplitPane verticalSP = new JSplitPane(JSplitPane.VERTICAL_SPLIT);
        verticalSP.setTopComponent(new JScrollPane(waybillPanel));
        verticalSP.setBottomComponent(horizontalSP);

        setLayout(new BorderLayout());
        add(toolBat, BorderLayout.NORTH);
        add(verticalSP, BorderLayout.CENTER);
    }

    public static void main(String[] args)
    {
        java.awt.EventQueue.invokeLater(() ->
        {
            Main main = new Main();
            main.setDefaultCloseOperation(WindowConstants.EXIT_ON_CLOSE);
            main.setVisible(true);
        });
    }

    public WaybillPanel getWaybillPanel()
    {
        return waybillPanel;
    }

    public RTree getTree()
    {
        return tree;
    }

    public RTable getTable()
    {
        return table;
    }

    public RToolBar getToolBat()
    {
        return toolBat;
    }

    public RMenuBar getRMenuBar()
    {
        return menuBar;
    }

    public SimpleDateFormat getDateFormat()
    {
        return dateFormat;
    }

    public SimpleDateFormat getDataBaseDateFormat()
    {
        return dataBaseDateFormat;
    }
}
