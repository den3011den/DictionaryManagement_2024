using System.Text.RegularExpressions;

public static class StringExtension
{
    public static string HideSlash(this string value)
    {
        return value.Replace("\\", "__123455432__");
    }

    public static string UnhideSlash(this string value)
    {
        return value.Replace("__123455432__", "\\");
    }

    public static string CleanNonPrintableSymbols(this string value)
    {
        return Regex.Replace(value, @"\p{C}+", string.Empty);
    }
    public static int Count(this string input, string substr)
    {
        int freq = 0;

        int index = input.IndexOf(substr);
        while (index >= 0)
        {
            index = input.IndexOf(substr, index + substr.Length);
            freq++;
        }
        return freq;
    }
}