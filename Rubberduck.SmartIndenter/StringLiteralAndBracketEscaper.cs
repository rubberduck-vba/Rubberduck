using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.SmartIndenter
{
    internal class StringLiteralAndBracketEscaper
    {
        public const char StringPlaceholder = '\a';
        public const char BracketPlaceholder = '\x02';

        private static readonly Regex StringReplaceRegex = new Regex("\a+", RegexOptions.IgnoreCase);
        private static readonly Regex BracketReplaceRegex = new Regex("\x02+", RegexOptions.IgnoreCase);

        private readonly List<string> _strings = new List<string>();
        private readonly List<string> _brackets = new List<string>();

        public string EscapedString { get; }

        public string OriginalString { get; }

        public IEnumerable<string> EscapedStrings => _strings;
        public IEnumerable<string> EscapedBrackets => _brackets;

        public string UnescapeIndented(string indented)
        {
            var code = indented;
            if (_strings.Any())
            {
                code = _strings.Aggregate(code, (current, literal) => StringReplaceRegex.Replace(current, literal, 1));
            }
            if (_brackets.Any())
            {
                code = _brackets.Aggregate(code, (current, expr) => BracketReplaceRegex.Replace(current, expr, 1));
            }
            return code;
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
                    if (c + 1 < chars.Length && chars[c] == '"')
                    {
                        c++;
                    }
                    quoted = false;
                    _strings.Add(new string(chars.Skip(strpos).Take(c - strpos).ToArray()));
                    for (var e = strpos; e < c; e++)
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
                    _brackets.Add(new string(chars.Skip(brkpos).Take(c - brkpos + 1).ToArray()));
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
