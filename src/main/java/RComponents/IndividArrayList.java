package RComponents;

import java.util.ArrayList;

public class IndividArrayList<String> extends ArrayList<String>
{
    public boolean add(String e)
    {
        boolean b = true;
        for(int i = 0; i < this.size(); i++)
        {
            if(this.get(i).equals(e))
                b = false;
        }
        if (b)
            super.add(e);

        return b;
    }
}
