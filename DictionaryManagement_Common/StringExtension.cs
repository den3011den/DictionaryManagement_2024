using System.Text;
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

}