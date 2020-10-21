using System;

namespace SharePoint.IO.Services
{
    internal class ContentService
    {
        public static string Process(string s, string[] defines)
        {
            var exclude = 0;
            var start = 0;
            var lineno = 1;
            var start_lineno = 1;
            for (var i = 0; i < s.Length; i++)
            {
                if (s[i] == '\n') lineno++;
                if (s[i] != '#' || (i > 0 && s[i - 1] != '\n')) continue;
                int removeCount;
                if (s.Substring(i, 6) == "#endif" && char.IsWhiteSpace(s[i + 6]))
                {
                    if (exclude > 0)
                    {
                        exclude--;
                        if (exclude == 0)
                        {
                            removeCount = i - start;
                            var v = s.Substring(start, removeCount);
                            var lfs = v.Length - v.Replace("\n", "").Length;
                            s = s.Remove(start, removeCount);
                            i -= removeCount;
                            s = s.Insert(start, new string('\n', lfs));
                            i += lfs;
                        }
                    }
                    removeCount = 0;
                    for (var j = i; j < s.Length && s[j] != '\n'; j++) removeCount++;
                    s = s.Remove(i, removeCount);
                }
                else if ((s.Substring(i, 6) == "#ifdef" && char.IsWhiteSpace(s[i + 6])) || (s.Substring(i, 7) == "#ifndef" && char.IsWhiteSpace(s[i + 7])))
                {
                    if (exclude > 0)
                        exclude++;
                    else
                    {
                        var j = i + 7;
                        for (; char.IsWhiteSpace(s[j]); j++) { }
                        var n = 0;
                        for (; j + n < s.Length && !char.IsWhiteSpace(s[j + n]); n++) { }
                        exclude = 1;
                        var flag = s.Substring(j, n);
                        if (defines != null)
                            for (var k = 0; k < defines.Length; k++)
                                if (defines[k] == flag)
                                {
                                    exclude = 0;
                                    break;
                                }
                        if (s[i + 3] == 'n') exclude = exclude == 0 ? 1 : 0;
                        if (exclude > 0)
                        {
                            start = i;
                            start_lineno = lineno;
                        }
                    }
                    removeCount = 0;
                    for (var j = i; j < s.Length && s[j] != '\n'; j++) removeCount++;
                    s = s.Remove(i, removeCount);
                }
            }
            if (exclude > 0)
                throw new Exception(string.Format("unterminated #ifdef starting on line {0}\n", start_lineno));
            return s;
        }
    }
}
