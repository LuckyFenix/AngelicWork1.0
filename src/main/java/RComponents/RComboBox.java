package RComponents;

import javax.swing.*;
import java.util.ArrayList;
import java.util.Collections;

public class RComboBox extends JComboBox
{
    public RComboBox(ArrayList<String> arrayList)
    {
        Collections.sort(arrayList);
        arrayList.forEach(this::addItem);
    }

    public RComboBox()
    {
        super();
    }

    protected RComboBox getThis()
    {
        return this;
    }

    public void changeComboBox(final ArrayList<String> arrayList)
    {
        Collections.sort(arrayList);
        removeAllItems();
        arrayList.forEach(this::addItem);
    }

    public boolean containString(String s)
    {
        for(int i = 0; i < getThis().getItemCount(); i++)
        {
            if(getThis().getItemAt(i).toString().equals(s))
                return true;
        }
        return false;
    }

    public ArrayList<String> getItemArray()
    {
        ArrayList<String> arrayList = new ArrayList<>();
        for(int i = 0; i < getThis().getItemCount(); i++)
        {
            arrayList.add(getThis().getItemAt(i).toString());
        }
        return arrayList;
    }
}
