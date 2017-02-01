using System.Collections.Generic;
using System.Linq;

namespace Rubberduck.SmartIndenter
{
    internal class StringLiteralAndBracketEscaper
    {
        public const char StringPlaceholder = '\a';
        public const char BracketPlaceholder = '\x2';

        private readonly List<string> _strings = new List<string>();
        private readonly List<string> _brackets = new List<string>();
        private readonly string _unescaped;
        private readonly string _escaped;

        public string EscapedString { get { return _escaped; } }
        public string OriginalString { get { return _unescaped; } } 
        public IEnumerable<string> EscapedStrings { get { return _strings; } }
        public IEnumerable<string> EscapedBrackets { get { return _brackets; } }

        public StringLiteralAndBracketEscaper(string code)
        {
            _unescaped = code;

            var chars = _unescaped.ToCharArray();
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
                    _brackets.Add(new string(chars.Skip(brkpos).Take(c - brkpos).ToArray()));
                    for (var e = brkpos; e < c; e++)
                    {
                        chars[e] = BracketPlaceholder;
                    }
                }
            }
            _escaped = new string(chars);
        }
    }
}
