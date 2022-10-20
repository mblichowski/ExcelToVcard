using System;

namespace ExcelToVcard;

public static class StringExtensions
{
    public static string FixChars(this object o) => o
        .ToString()?
        .Replace(",", " ")
        .Replace(Environment.NewLine, " ")
        .Replace("\n", " ")
        .Trim() ?? String.Empty;

    public static string FixUrl(this string s)
    {
        if (s.EndsWith("/"))
            s = s[..^1];

        return s
            .Replace("http://", "")
            .Replace("https://", "");
    }

    public static string CheckPhone(this string s)
    {
        return
            s.StartsWith("+") ?
            s :
            "+" + s;
    }
}
