package RComponents;

import javax.swing.*;
import javax.swing.event.DocumentEvent;
import javax.swing.event.DocumentListener;
import javax.swing.text.JTextComponent;
import java.util.ArrayList;
import java.util.Collections;

public class LiveSearchComboBox extends JComboBox
{
    private ArrayList<String> arrayList = new ArrayList<>();
    private ArrayList<String> arrayRemoveItem = new ArrayList<>();
    protected String inputText;
    protected boolean bool = true;

    public LiveSearchComboBox(ArrayList<String> arrayList)
    {
        Collections.sort(arrayList);
        arrayList.forEach(this::addItem);

        setEditable(true);
        ((JTextComponent) getEditor().getEditorComponent()).setText("");
        ((JTextComponent) getEditor().getEditorComponent()).getDocument().addDocumentListener(new ComboBoxListener());

        this.arrayList = arrayList;
    }

    protected LiveSearchComboBox getThis()
    {
        return this;
    }

    private class ComboBoxListener implements DocumentListener
    {
        @Override
        public void insertUpdate(DocumentEvent e)
        {
            if(bool)
            {
                getThis().setPopupVisible(true);
                inputText = ((JTextComponent) getThis().getEditor().getEditorComponent()).getText();
                liveSearch();
            }
        }

        @Override
        public void removeUpdate(DocumentEvent e)
        {
            if(bool)
            {
                getThis().setPopupVisible(true);
                inputText = ((JTextComponent) getThis().getEditor().getEditorComponent()).getText();
                liveSearch();
            }
        }

        @Override
        public void changedUpdate(DocumentEvent e)
        {
        }
    }

    public void liveSearch()
    {
        if (inputText.length() != 0)
        {
            for(int i = 0; i < arrayList.size(); i++)
            {
                if(!arrayList.get(i).contains(inputText))
                {
                    arrayRemoveItem.add(arrayList.get(i));
                    arrayList.remove(i);
                    i--;
                }
            }
            for(int i = 0; i < arrayRemoveItem.size(); i++)
            {
                if(arrayRemoveItem.get(i).contains(inputText))
                {
                    arrayList.add(arrayRemoveItem.get(i));
                    arrayRemoveItem.remove(i);
                    i--;
                }
            }
            if(arrayList.isEmpty())
                Collections.sort(arrayList);
            bool = false;
            SwingUtilities.invokeLater(() ->
            {
                getThis().removeAllItems();
                for (String anArrayList : arrayList)
                {
                    getThis().addItem(anArrayList);
                }
                ((JTextComponent) getThis().getEditor().getEditorComponent()).setText(inputText);
                bool = true;
            });
        } else
            getThis().setPopupVisible(false);
    }

    public void changeComboBox(final ArrayList<String> arrayList)
    {
        SwingUtilities.invokeLater(() ->
        {
            Collections.sort(arrayList);
            removeAllItems();
            arrayList.forEach(this::addItem);
            ((JTextField) getThis().getEditor().getEditorComponent()).setText("");
            getThis().setPopupVisible(false);
            getThis().arrayList = arrayList;
        });
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
}
