package Support;

import java.util.HashMap;
import java.util.Map.Entry;

public class TranslatingMap extends HashMap<String, String>
{
    public String getValue(String key)
    {
        String value = key;

        for( Entry<String, String> entry : this.entrySet())
            if(key.equals(entry.getKey()))
                value = entry.getValue();

        return value;
    }

    public String getKey(String value)
    {
        String key = value;

        for( Entry<String, String> entry : this.entrySet())
            if(value.equals(entry.getValue()))
                key = entry.getKey();

        return key;
    }
}
