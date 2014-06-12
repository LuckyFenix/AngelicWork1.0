package RDialogs;

import Support.DBI;
import Support.GBC;
import Main.Main;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

import javax.swing.*;
import java.awt.*;
import java.awt.event.ActionEvent;
import java.awt.event.ActionListener;
import java.io.File;
import java.sql.SQLException;
import java.util.ArrayList;

public class ImportDialog extends RDialog
{
    private Main main;
    private final JLabel label;
    private String selectedFile;
    private final JTextField startNum;
    private final JTextField endNum;
    private final ArrayList columns;
    private final JComboBox category;
    private final JComboBox group;
    private final JComboBox name;
    private final JComboBox unitOfMeasure;
    private final JComboBox<Object> purchasePrice;
    private final JComboBox<Object> warehouseResidue;
    private final JComboBox shopResidue;
    private final JComboBox sellingPrice;

    public ImportDialog(Main main)
    {
        super(main, "Імпорт залишків", true);
        this.main = main;

        GridBagLayout layout = new GridBagLayout();
        setLayout(layout);

        columns = convertArray(new String[]{"A", "B", "C", "D", "E", "F", "G", "H", "I", "G", "K"});
        category = new JComboBox(columns.toArray());
        category.setSelectedItem("A");
        group = new JComboBox(columns.toArray());
        group.setSelectedItem("B");
        name = new JComboBox(columns.toArray());
        name.setSelectedItem("C");
        unitOfMeasure = new JComboBox(columns.toArray());
        unitOfMeasure.setSelectedItem("D");
        warehouseResidue = new JComboBox<Object>(columns.toArray());
        warehouseResidue.setSelectedItem("E");
        shopResidue = new JComboBox(columns.toArray());
        shopResidue.setSelectedItem("F");
        sellingPrice = new JComboBox(columns.toArray());
        sellingPrice.setSelectedItem("G");
        purchasePrice = new JComboBox<>(columns.toArray());
        purchasePrice.setSelectedItem("H");


        startNum = new JTextField(5);
        endNum = new JTextField(5);

        JButton fileBtn = new JButton("Файл");
        JButton startBtn = new JButton("Старт");
        JButton closeBtn = new JButton("Закрити");

        String s = "Вибраний документ: Документ не вибрано";
        label = new JLabel(s);
        Font f = new Font(label.getFont().getName(), Font.BOLD, 13);
        label.setFont(f);

        add(label, new GBC(0, 0, 3, 1).setAnchor(GBC.WEST).setInsets(20));
        add(new JLabel("Категорія товару:"), new GBC(0, 1, 2, 1).setAnchor(GBC.WEST));
        add(category, new GBC(2, 1, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Группа товару:"), new GBC(0, 2, 2, 1).setAnchor(GBC.WEST));
        add(group, new GBC(2, 2, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Назва товару:"), new GBC(0, 3, 2, 1).setAnchor(GBC.WEST));
        add(name, new GBC(2, 3, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Одиниці вимірювання:"), new GBC(0, 4, 2, 1).setAnchor(GBC.WEST));
        add(unitOfMeasure, new GBC(2, 4, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Залишок на складі:"), new GBC(0, 5, 2, 1).setAnchor(GBC.WEST));
        add(warehouseResidue, new GBC(2, 5, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Залишок в магазині:"), new GBC(0, 6, 2, 1).setAnchor(GBC.WEST));
        add(shopResidue, new GBC(2, 6, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Продажна ціна:"), new GBC(0, 7, 2, 1).setAnchor(GBC.WEST));
        add(sellingPrice, new GBC(2, 7, 1, 1).setAnchor(GBC.EAST));
        add(new JLabel("Прихідна ціна:"), new GBC(0, 8, 2, 1).setAnchor(GBC.WEST));
        add(purchasePrice, new GBC(2, 8, 1, 1).setAnchor(GBC.EAST));

        JPanel linesNumPanel = new JPanel(new GridBagLayout());
        linesNumPanel.add(new JLabel("Читати рядки з "), new GBC(0, 0));
        linesNumPanel.add(startNum, new GBC(1, 0));
        linesNumPanel.add(new JLabel(" по "), new GBC(2, 0));
        linesNumPanel.add(endNum, new GBC(3, 0));
        linesNumPanel.add(new JLabel("; "), new GBC(4, 0));
        JLabel iteration = new JLabel("");
        linesNumPanel.add(iteration, new GBC(5, 0));
        add(linesNumPanel, new GBC(0, 9, 3, 1).setAnchor(GBC.WEST));

        JPanel btnPanel = new JPanel(new GridBagLayout());
        btnPanel.add(fileBtn, new GBC(0, 0).setAnchor(GBC.CENTER).setInsets(0, 25, 0, 25));
        btnPanel.add(startBtn, new GBC(1, 0).setAnchor(GBC.CENTER).setInsets(0, 25, 0, 25));
        btnPanel.add(closeBtn, new GBC(2, 0).setAnchor(GBC.CENTER).setInsets(0, 25, 0, 25));
        add(btnPanel, new GBC(0, 10, 3, 1).setAnchor(GBC.CENTER).setInsets(10));

        fileBtn.addActionListener(new fileBtnList());
        startBtn.addActionListener(new startBtnList());
        closeBtn.addActionListener(new closeBtnAction());

        pack();
        setLocationRelativeTo(main);
        setLocation(main.getX() + main.getWidth()/2 - this.getWidth()/2, main.getY() + main.getHeight()/2 - this.getHeight()/2);
        setVisible(true);
    }

    private ImportDialog getThisFrame()
    {
        return this;
    }



    private class fileBtnList implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            JFileChooser fileChooser = new JFileChooser(new File("."));
            int result = fileChooser.showOpenDialog(getThisFrame());
            if(result == JFileChooser.APPROVE_OPTION)
            {
                selectedFile = fileChooser.getSelectedFile().getPath();
                label.setText("Вибраний документ: " + selectedFile);
                getThisFrame().pack();
            }
        }
    }

    int i;
    private class startBtnList implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            try
            {
                final Thread thread = new Thread(new Runnable()
                {
                    private long startT;
                    private ArrayList rowArray = new ArrayList<String>();
                    int startRow;
                    int endRow;
                    DBI warehouseDBI;
                    HSSFWorkbook wb = readWorkbook(selectedFile);
                    HSSFSheet sheet = wb.getSheetAt(0);
                    HSSFRow row;
                    String categoryName;
                    String groupName;
                    String nameName;
                    String unitOfMeasureName;
                    double purchasePriceValue;
                    double warehouseResidueValue;
                    double shopResidueValue;
                    double totalItems;
                    double sellingPriceValue;
                    @Override
                    public void run()
                    {
                        startRow = Integer.parseInt(startNum.getText());
                        endRow = Integer.parseInt(endNum.getText());
                        startT = System.currentTimeMillis();
                        getThisFrame().setSize(new Dimension(getThisFrame().getWidth() + 500, getThisFrame().getHeight()));
                        for(i = startRow - 1; i<endRow; i++)
                        {
                            try
                            {
                                getThisFrame().setTitle("Запис№" + (i + 1) + "; " +
                                        "Часу залишилося: " + ((int)
                                        ((((System.currentTimeMillis() - startT)
                                                * (endRow - startRow + 1)
                                                /  (i - startRow + 1))
                                                - (System.currentTimeMillis() - startT))
                                        / 1000))
                                         +  " сек;" +
                                        "Часу пройшло:" + ((System.currentTimeMillis() - startT) / 1000) + "сек;" +
                                        "Часу всього необхідно" + ((System.currentTimeMillis() - startT) * (endRow - startRow + 1) / (i - startRow + 1) / 1000) + "сек;");
                            } catch (ArithmeticException ignored) {}

                            row = sheet.getRow(i);

                            rowArray.clear();
                            categoryName = smartTrim(row.getCell(columns.indexOf(category.getSelectedItem().toString())).getStringCellValue()); rowArray.add(categoryName);
                            groupName = smartTrim(row.getCell(columns.indexOf(group.getSelectedItem().toString())).getStringCellValue()); rowArray.add(groupName);
                            nameName = smartTrim(row.getCell(columns.indexOf(name.getSelectedItem().toString())).getStringCellValue()); rowArray.add(nameName);
                            unitOfMeasureName = row.getCell(columns.indexOf(unitOfMeasure.getSelectedItem().toString())).getStringCellValue(); rowArray.add(unitOfMeasureName);
                            warehouseResidueValue = row.getCell(columns.indexOf(warehouseResidue.getSelectedItem().toString())).getNumericCellValue(); rowArray.add(warehouseResidueValue);
                            shopResidueValue = row.getCell(columns.indexOf(shopResidue.getSelectedItem().toString())).getNumericCellValue(); rowArray.add(shopResidueValue);
                            totalItems = warehouseResidueValue + shopResidueValue; rowArray.add(totalItems);
                            sellingPriceValue = row.getCell(columns.indexOf(sellingPrice.getSelectedItem().toString())).getNumericCellValue(); rowArray.add(sellingPriceValue);
                            purchasePriceValue = row.getCell(columns.indexOf(purchasePrice.getSelectedItem().toString())).getNumericCellValue(); rowArray.add(purchasePriceValue);

                            main.getTree().createGroup(categoryName, groupName);
                            warehouseDBI = new DBI("databassesabc");
                            try
                            {
                                main.getTree().createGroup(categoryName, groupName);
                                warehouseDBI.getSt().execute("INSERT " +
                                        "INTO `товари` (`Категорія`, `Група`, `Назва товару`, `Одиниці вимірювання`, `Залишок на початок. Склад`, `Залишок на початок. Магазин`, `Товару в загальному`, `Ціна`, `Ціна закупочна`, `Дата останнього приходу`, `Бронь`) " +
                                        "VALUE (" + getStringFromArray(rowArray, '\'') + ", '0000-00-00', '0');");
                            } catch (SQLException e1)
                            {
                                System.err.println(categoryName + " " + groupName + " " + nameName);
                                e1.printStackTrace();
                            }
                            warehouseDBI.close(main.k);
                        }
                        main.getTree().updatePanel();
                        getThisFrame().setTitle("Імпорт залишків");
                        getThisFrame().setSize(new Dimension(getThisFrame().getWidth() - 500, getThisFrame().getHeight()));
                        JOptionPane.showMessageDialog(getThisFrame(), "Імпорт завершено! Імпортовано " + (i - startRow + 1) + " од. товару.");
                    }
                });
                thread.start();
            } catch (Exception e1)
            {
                e1.printStackTrace();
            }
        }
    }

    private class closeBtnAction implements ActionListener
    {
        @Override
        public void actionPerformed(ActionEvent e)
        {
            getThisFrame().dispose();
        }
    }


}
