package RComponents;

import Main.Main;
import RDialogs.OpenWaybillDialog;
import RDialogs.SuppliersDialog;

import javax.swing.*;


public class RToolBar extends JToolBar
{
    private final Main main;
    private final JComboBox<String> comboBox;

    public RToolBar(final Main main)
    {
        this.main = main;

        JButton newDoc = new JButton(new ImageIcon(getClass().getResource("/icon/file_new.png")));
        JButton openDoc = new JButton(new ImageIcon(getClass().getResource("/icon/file_open.png")));
        JButton saveDoc = new JButton(new ImageIcon(getClass().getResource("/icon/file_save.png")));
        JButton deleteDoc = new JButton(new ImageIcon(getClass().getResource("/icon/file_delete.png")));
        JButton reportDoc = new JButton(new ImageIcon(getClass().getResource("/icon/file_report.png")));
        JButton suppliers = new JButton(new ImageIcon(getClass().getResource("/icon/file_supplier.png")));

        String[] s = new String[]{"Прихідна накладна", "Відпускна накладна", "Ревізія"};
        comboBox = new JComboBox<>(s);

        newDoc.addActionListener(e ->
        {
            if (comboBox.getSelectedItem().toString().equals("Ревізія"))
                main.getWaybillPanel().getRevision().initNewRevision();
            else
                main.getWaybillPanel().init(comboBox.getSelectedItem().toString());
        });

        openDoc.addActionListener(e -> new OpenWaybillDialog(main, comboBox.getSelectedItem().toString()));

        suppliers.addActionListener(e ->
        {
            new SuppliersDialog(main);
        });

        comboBox.addActionListener(e ->
        {
            main.getRMenuBar().getWaybillMenu().setText(comboBox.getSelectedItem().toString());
            main.getWaybillPanel().init(comboBox.getSelectedItem().toString());
        });

        add(newDoc);
        add(openDoc);
        add(saveDoc);
        add(deleteDoc);
        add(reportDoc);

        addSeparator();

        add(suppliers);

        addSeparator();
        add(comboBox);

        addSeparator();
    }

    public JComboBox<String> getComboBox()
    {
        return comboBox;
    }
}
