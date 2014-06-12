package RComponents;

import Main.Main;
import Support.GBC;
import Waybills.Invoice;
import Waybills.Revision;
import Waybills.WaybillSelling;

import javax.swing.*;
import java.awt.*;

public class WaybillPanel extends JPanel
{
    private Main main;
    private Invoice invoice;
    private Revision revision;
    private WaybillSelling waybillSelling;

    public WaybillPanel(Main main)
    {
        this.main = main;

        this.removeAll();
        setLayout(new GridBagLayout());

        init(main.getToolBat().getComboBox().getSelectedItem().toString());
    }

    public void init(String waybill)
    {
        removeAll();
        repaint();

        if (waybill.equals("Прихідна накладна"))
        {
            invoice = new Invoice(main);
            add(invoice, new GBC(0, 0).setAnchor(GBC.NORTHWEST));
        }
        if (waybill.equals("Відпускна накладна"))
        {
            waybillSelling = new WaybillSelling(main);
            add(waybillSelling, new GBC(0, 0).setAnchor(GBC.NORTHWEST));
        }
        if (waybill.equals("Ревізія"))
        {
            revision = new Revision(main);
            add(revision, new GBC(0, 0).setAnchor(GBC.NORTHWEST));
        }
        revalidate();
    }

    public Invoice getInvoice()
    {
        return invoice;
    }

    public Revision getRevision()
    {
        return revision;
    }

    public WaybillSelling getWaybillSelling()
    {
        return waybillSelling;
    }
}
