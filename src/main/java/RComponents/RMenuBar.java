package RComponents;

import Main.Main;
import RDialogs.ExportDialog;
import RDialogs.ExportToReimportDialog;
import RDialogs.ImportDialog;
import RDialogs.StaffDialog;

import javax.swing.*;
import java.sql.SQLException;


public class RMenuBar extends JMenuBar
{
    Main main;
    private final JMenu waybillMenu;

    public RMenuBar(final Main main)
    {
        this.main = main;
        waybillMenu = new JMenu(main.getToolBat().getComboBox().getSelectedItem().toString());

        JMenu warehouseMenu = new JMenu("Склад");

        JMenu suppliersMenu = new JMenu("Постачальники");

        JMenu staffMenu = new JMenu("Працівники");
        JMenuItem staffEdit = new JMenuItem("Меню працівників");
        staffEdit.addActionListener(e ->
        {
            new StaffDialog(main);
        });
        staffMenu.add(staffEdit);

        JMenu reportsMenu = new JMenu("Звіти");

        JMenu serviceMenu = new JMenu("Сервіс");
        JMenuItem remeainsImport = new JMenuItem("Імпорт залишків з Excel документа");
        JMenuItem remeainsExportToRevision = new JMenuItem("Експорт залишків в Excel для ревізії");
        JMenuItem remeainsExportToReimport = new JMenuItem("Експорт залишків в Excel для реімпорту");
        remeainsImport.addActionListener(e -> new ImportDialog(main));
        remeainsExportToRevision.addActionListener(e ->
        {
            try
            {
                new ExportDialog(main);
            } catch (SQLException e1) {e1.printStackTrace();}
        });
        remeainsExportToReimport.addActionListener(e ->
        {
            try
            {
                new ExportToReimportDialog(main);
            } catch (SQLException e1)
            {
                e1.printStackTrace();
            }
        });
        serviceMenu.add(remeainsImport);
        serviceMenu.add(remeainsExportToRevision);
        serviceMenu.add(remeainsExportToReimport);

        add(waybillMenu);
        add(warehouseMenu);
        add(suppliersMenu);
        add(staffMenu);
        add(reportsMenu);
        add(serviceMenu);
    }

    public JMenu getWaybillMenu()
    {
        return waybillMenu;
    }
}
