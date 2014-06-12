package Support;

import Support.TranslatingMap;

public class Translator
{
    private TranslatingMap translator = new TranslatingMap();

    public Translator()
    {
        translator.put("а", "a");
        translator.put("б", "b");
        translator.put("в", "v");
        translator.put("г", "g");
        translator.put("д", "d");
        translator.put("е", "e");
        translator.put("є", "3");
        translator.put("ж", "j");
        translator.put("з", "z");
        translator.put("и", "u");
        translator.put("і", "i");
        translator.put("ї", "1");
        translator.put("й", "y");
        translator.put("к", "k");
        translator.put("л", "l");
        translator.put("м", "m");
        translator.put("н", "n");
        translator.put("о", "o");
        translator.put("п", "p");
        translator.put("р", "r");
        translator.put("с", "s");
        translator.put("т", "t");
        translator.put("у", "q");
        translator.put("ф", "f");
        translator.put("х", "h");
        translator.put("ц", "c");
        translator.put("ч", "4");
        translator.put("ш", "w");
        translator.put("щ", "8");
        translator.put("ь", "6");
        translator.put("ю", "0");
        translator.put("я", "9");
        translator.put(" ", "_");
        translator.put("А", "A");
        translator.put("Б", "B");
        translator.put("В", "V");
        translator.put("Г", "G");
        translator.put("Д", "D");
        translator.put("Е", "E");
        translator.put("Є", "2");
        translator.put("Ж", "J");
        translator.put("З", "Z");
        translator.put("И", "U");
        translator.put("І", "I");
        translator.put("Ї", "5");
        translator.put("Й", "Y");
        translator.put("К", "K");
        translator.put("Л", "L");
        translator.put("М", "M");
        translator.put("Н", "N");
        translator.put("О", "O");
        translator.put("П", "P");
        translator.put("Р", "R");
        translator.put("С", "S");
        translator.put("Т", "T");
        translator.put("У", "Q");
        translator.put("Ф", "F");
        translator.put("Х", "H");
        translator.put("Ц", "C");
        translator.put("Ч", "7");
        translator.put("Ш", "W");
        translator.put("Щ", "~");
        translator.put("Ь", "+");
        translator.put("Ю", "-");
        translator.put("Я", "@");
    }

    public String translateToEng(String ukrVord)
    {
        StringBuilder buffer = new StringBuilder();
        buffer.setLength(ukrVord.length());
        for(int i = 0; i<ukrVord.length(); i++)
        {
            char ch = translator.getValue("" + ukrVord.charAt(i)).charAt(0);
            buffer.setCharAt(i, ch);
        }

        return buffer.substring(0);
    }

    public String translateToUkr(String engVord)
    {
        StringBuilder buffer = new StringBuilder();
        buffer.setLength(engVord.length());
        for(int i = 0; i<engVord.length(); i++)
        {
            buffer.setCharAt(i, translator.getValue("" + engVord.charAt(i)).charAt(0));
        }
        return buffer.substring(0);
    }
}
