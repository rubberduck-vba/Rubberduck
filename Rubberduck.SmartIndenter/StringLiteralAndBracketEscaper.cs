using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.SmartIndenter
{
    internal class StringLiteralAndBracketEscaper
    {
        public const char StringPlaceholder = '\a';
        public const char BracketPlaceholder = '\x02';

        private readonly List<string> _strings = new List<string>();
        private readonly List<string> _brackets = new List<string>();

        public string EscapedString { get; }

        public string OriginalString { get; }

        public IEnumerable<string> EscapedStrings => _strings;
        public IEnumerable<string> EscapedBrackets => _brackets;

        public string UnescapeIndented(string indented)
        {

            var code = ReplaceEscapedItems(indented, StringPlaceholder, EscapedStrings);
            return ReplaceEscapedItems(code, BracketPlaceholder, EscapedBrackets);
        }

        private string ReplaceEscapedItems(string code, char placehoder, IEnumerable<string> replacements)
        {
            var output = code;
            foreach (var item in replacements)
            {
                var pos = output.IndexOf(new string(placehoder, item.Length), StringComparison.Ordinal);
                output = output.Substring(0, pos) + item + output.Substring(pos + item.Length);
            }
            return output;
        }

        public StringLiteralAndBracketEscaper(string code)
        {
            OriginalString = code;

            var chars = OriginalString.ToCharArray();
            var quoted = false;
            var bracketed = false;
            var ins = 0;
            var strpos = 0;
            var brkpos = 0;
            for (var c = 0; c < chars.Length; c++)
            {
                if (chars[c] == '"' && !bracketed)
                {
                    if (!quoted)
                    {
                        strpos = c;
                        quoted = true;
                        continue;
                    }
                    if (c + 1 < chars.Length && chars[c + 1] == '"')
                    {
                        c++;
                    }
                    quoted = false;
                    _strings.Add(OriginalString.Substring(strpos, c - strpos + 1));
                    for (var e = strpos; e <= c; e++)
                    {
                        chars[e] = StringPlaceholder;
                    }
                }
                else if (!quoted && !bracketed && chars[c] == '[')
                {
                    bracketed = true;
                    brkpos = c;
                    ins++;
                }
                else if (!quoted && bracketed && chars[c] == ']')
                {
                    ins--;
                    if (ins != 0)
                    {
                        continue;
                    }
                    bracketed = false;
                    _brackets.Add(OriginalString.Substring(brkpos, c - brkpos + 1));
                    for (var e = brkpos; e <= c; e++)
                    {
                        chars[e] = BracketPlaceholder;
                    }
                }
            }
            EscapedString = new string(chars);
        }
    }
}
