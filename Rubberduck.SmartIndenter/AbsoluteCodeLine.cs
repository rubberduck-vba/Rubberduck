using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.SmartIndenter
{
    internal class AbsoluteCodeLine
    {
        private const string StupidLineEnding = ": _";
        private const char StringPlaceholder = '\a';
        private const char BracketPlaceholder = '\x2';
        private static readonly Regex StringReplaceRegex = new Regex(StringPlaceholder.ToString(CultureInfo.InvariantCulture));
        private static readonly Regex BracketReplaceRegex = new Regex(BracketPlaceholder.ToString(CultureInfo.InvariantCulture));
        private static readonly Regex LineNumberRegex = new Regex(@"^(?<number>(-?\d+)|(&H[0-9A-F]{1,8}))\s+(?<code>.*)", RegexOptions.ExplicitCapture);
        private static readonly Regex EndOfLineCommentRegex = new Regex(@"^(?!(Rem\s)|('))(?<code>[^']*)(\s(?<comment>'.*))$", RegexOptions.ExplicitCapture);      
        private static readonly Regex ProcedureStartRegex = new Regex(@"^(Public\s|Private\s|Friend\s)?(Static\s)?(Sub|Function|Property\s(Let|Get|Set))\s");
        private static readonly Regex ProcedureStartIgnoreRegex = new Regex(@"^[LR]?Set\s|^Let\s|^(Public|Private)\sDeclare\s(Function|Sub)");
        private static readonly Regex ProcedureEndRegex = new Regex(@"^End\s(Sub|Function|Property)");
        private static readonly Regex TypeEnumStartRegex = new Regex(@"^(Public\s|Private\s)?(Enum\s|Type\s)");
        private static readonly Regex TypeEnumEndRegex = new Regex(@"^End\s(Enum|Type)");
        private static readonly Regex InProcedureInRegex = new Regex(@"^(Else)?If\s.*\sThen$|^Else$|^Case\s|^With|^For\s|^Do$|^Do\s|^While$|^While\s|^Select Case");
        private static readonly Regex InProcedureOutRegex = new Regex(@"^Else(If)?|^Case\s|^End With|^Next\s|^Next$|^Loop$|^Loop\s|^Wend$|^End If$|^End Select");
        private static readonly Regex DeclarationRegex = new Regex(@"^(Dim|Const|Static|Public|Private)\s(.*(\sAs\s)?|_)");
        private static readonly Regex PrecompilerInRegex = new Regex(@"^#(Else)?If\s.+Then$|^#Else$");
        private static readonly Regex PrecompilerOutRegex = new Regex(@"^#ElseIf\s.+Then|^#Else$|#End\sIf$");
        private static readonly Regex SingleLineElseIfRegex = new Regex(@"^ElseIf\s.*\sThen\s.*");

        private readonly IIndenterSettings _settings;
        private int _lineNumber;
        private bool _numbered;
        private string _code;
        private readonly bool _stupidLineEnding;
        private readonly string[] _segments;
        private List<string> _strings;
        private List<string> _brackets;

        public AbsoluteCodeLine(string code, IIndenterSettings settings) : this(code, settings, null) { }

        public AbsoluteCodeLine(string code, IIndenterSettings settings, AbsoluteCodeLine previous)
        {
            _settings = settings;
            Previous = previous;

            if (code.EndsWith(StupidLineEnding))
            {
                _code = code.Substring(0, code.Length - StupidLineEnding.Length);
                _stupidLineEnding = true;
            }
            else
            {
                _code = code;
            }
            
            Original = code;

            ExtractStringLiteralsAndBrackets();
            ExtractLineNumber();
            ExtractEndOfLineComment();

            _code = Regex.Replace(_code, StringPlaceholder + "+", StringPlaceholder.ToString(CultureInfo.InvariantCulture));
            _code = Regex.Replace(_code, BracketPlaceholder + "+", BracketPlaceholder.ToString(CultureInfo.InvariantCulture)).Trim();
            _segments = _code.Split(new[] { ": " }, StringSplitOptions.None);
        }

        //TODO: This should be a class.
        private void ExtractStringLiteralsAndBrackets()
        {
            _strings = new List<string>();
            _brackets = new List<string>();

            var chars = _code.ToCharArray();
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
            _code = new string(chars);            
        }

        private void ExtractLineNumber()
        {
            if (Previous == null || !Previous.HasContinuation)
            {
                var match = LineNumberRegex.Match(_code);
                if (match.Success)
                {
                    _code = match.Groups["code"].Value;
                    _numbered = true;
                    var number = match.Groups["number"].Value;
                    if (!int.TryParse(number, out _lineNumber))
                    {
                        int.TryParse(number.Replace("&H", string.Empty), NumberStyles.HexNumber, 
                                     CultureInfo.InvariantCulture, out _lineNumber);
                    }                  
                }
            }
            _code = _code.Trim();
        }

        private void ExtractEndOfLineComment()
        {
            var match = EndOfLineCommentRegex.Match(_code);
            if (!match.Success)
            {
                EndOfLineComment = string.Empty;
                return;
            }
            _code = match.Groups["code"].Value.Trim();
            EndOfLineComment = match.Groups["comment"].Value.Trim();
        }

        public AbsoluteCodeLine Previous { get; private set; }

        public string Original { get; private set; }

        public string Escaped
        {
            get
            {
                var output = Original;
                foreach (var item in _strings)
                {
                    output = output.Replace(item, new string(StringPlaceholder, item.Length));
                }
                foreach (var item in _brackets)
                {
                    output = output.Replace(item, new string(BracketPlaceholder, item.Length));
                }
                return output;
            }
        }

        public string EndOfLineComment { get; private set; }

        public IEnumerable<string> Segments
        {
            get { return _segments; }
        }

        public string ContinuationRebuildText
        {
            get
            {
                var output = (_code + " " + EndOfLineComment).Trim();
                return HasContinuation ? output.Substring(0, output.Length - 1) : output;
            }
        }

        public bool ContainsOnlyComment
        {
            get { return _code.StartsWith("'") || _code.StartsWith("Rem "); }
        }

        public bool IsDeclaration
        {
            get { return !IsEmpty && (!IsProcedureStart && !ProcedureStartIgnoreRegex.IsMatch(_code)) && DeclarationRegex.IsMatch(_code); }
        }

        public bool IsDeclarationContinuation { get; set; }

        public bool HasDeclarationContinuation
        {
            get
            {
                return (!IsProcedureStart && !ProcedureStartIgnoreRegex.IsMatch(_code)) &&
                       !ContainsOnlyComment &&
                       string.IsNullOrEmpty(EndOfLineComment) &&
                       HasContinuation &&
                       ((IsDeclarationContinuation && Segments.Count() == 1) || DeclarationRegex.IsMatch(Segments.Last()));
            }
        }

        public bool HasContinuation
        {
            get { return _code.Equals("_") || _code.EndsWith(" _") || EndOfLineComment.EndsWith(" _"); }
        }

        public bool IsPrecompilerDirective
        {
            get { return _code.TrimStart().StartsWith("#"); }
        }

        public bool IsBareDebugStatement
        {
            get { return _code.StartsWith("Debug.") || _code.Equals("Stop"); }
        }

        public int EnumOrTypeStarts
        {
            get { return _segments.Count(s => TypeEnumStartRegex.IsMatch(s)); }
        }

        public int EnumOrTypeEnds
        {
            get { return _segments.Count(s => TypeEnumEndRegex.IsMatch(s)); }
        }

        public bool IsProcedureStart
        {
            get
            { return _segments.Any(s => ProcedureStartRegex.IsMatch(s)) && !_segments.Any(s => ProcedureStartIgnoreRegex.IsMatch(s)); }
        }

        public bool IsProcedureEnd
        {
            get { return _segments.Any(s => ProcedureEndRegex.IsMatch(s)); }
        }

        public bool IsEmpty
        {
            get { return Original.Trim().Length == 0; }
        }

        public int NextLineIndents
        {
            get
            {
                var adjust = _settings.IndentCase && _segments.Any(s => s.TrimStart().StartsWith("Select Case")) ? 1 : 0;
                var ins = _segments.Count(s => InProcedureInRegex.IsMatch(s.Trim())) + (IsProcedureStart && _settings.IndentEntireProcedureBody ? 1 : 0) + adjust;

                ins += _segments.Count(s => SingleLineElseIfRegex.IsMatch(s));
                ins -= MultipleCaseAdjustment();

                if (_settings.IndentCompilerDirectives && PrecompilerInRegex.IsMatch(_segments[0].Trim()))
                {
                    ins += 1;
                }
                return ins - Outdents;
            }
        }

        public int Outdents
        {
            get
            {
                var adjust = _settings.IndentCase && _segments.Any(s => s.TrimStart().StartsWith("End Select")) ? 1 : 0;
                var outs = _segments.Count(s => InProcedureOutRegex.IsMatch(s.Trim())) + (IsProcedureEnd && _settings.IndentEntireProcedureBody ? 1 : 0) + adjust;

                outs -= MultipleCaseAdjustment();

                if (_settings.IndentCompilerDirectives && PrecompilerOutRegex.IsMatch(_segments[0].Trim()))
                {
                    outs += 1;
                }
                return outs;
            }
        }

        private int MultipleCaseAdjustment()
        {
            var cases = _segments.Count(s => s.StartsWith("Case "));
            return (cases > 1 && _segments.Length % 2 != 0) ? cases - 1 : 0;
        }

        public string Indent(int indents, bool atProcStart, bool absolute = false)
        {
            if (IsEmpty || (ContainsOnlyComment && !_settings.AlignCommentsWithCode && !absolute))
            {
                return Original;
            }

            if ((IsPrecompilerDirective && _settings.ForceCompilerDirectivesInColumn1) ||
                (IsBareDebugStatement && _settings.ForceDebugStatementsInColumn1) ||
                (atProcStart && !_settings.IndentFirstCommentBlock && ContainsOnlyComment) ||
                (atProcStart && !_settings.IndentFirstDeclarationBlock && (IsDeclaration || IsDeclarationContinuation)))
            {
                indents = 0;
            }

            var number = _numbered ? _lineNumber.ToString(CultureInfo.InvariantCulture) : string.Empty;
            var gap = Math.Max((absolute ? indents : _settings.IndentSpaces * indents) - number.Length, number.Length > 0 ? 1 : 0);
            if (_settings.AlignDims && (IsDeclaration || IsDeclarationContinuation))
            {
                AlignDims(gap);
            }

            var code = string.Join(": ", _segments);
            if (_strings.Any())
            {
                code = _strings.Aggregate(code, (current, literal) => StringReplaceRegex.Replace(current, literal, 1));
            }
            if (_brackets.Any())
            {
                code = _brackets.Aggregate(code, (current, expr) => BracketReplaceRegex.Replace(current, expr, 1));
            }

            code = string.Join(string.Empty, number, new string(' ', gap), code);
            if (string.IsNullOrEmpty(EndOfLineComment))
            {
                return code + (_stupidLineEnding ? StupidLineEnding : string.Empty);
            }

            var position = Original.LastIndexOf(EndOfLineComment, StringComparison.Ordinal);
            switch (_settings.EndOfLineCommentStyle)
            {
                case EndOfLineCommentStyle.Absolute:
                    return string.Format("{0}{1}{2}{3}", code, new string(' ', Math.Max(position - code.Length, 1)),
                                                         EndOfLineComment, _stupidLineEnding ? StupidLineEnding : string.Empty);
                case EndOfLineCommentStyle.SameGap:
                    var uncommented = Original.Substring(0, position - 1);
                    return string.Format("{0}{1}{2}{3}", code, new string(' ', uncommented.Length - uncommented.TrimEnd().Length + 1), 
                                                         EndOfLineComment, _stupidLineEnding ? StupidLineEnding : string.Empty);
                case EndOfLineCommentStyle.StandardGap:
                    return string.Format("{0}{1}{2}{3}", code, new string(' ', _settings.IndentSpaces * 2), EndOfLineComment,
                                                        _stupidLineEnding ? StupidLineEnding : string.Empty);
                case EndOfLineCommentStyle.AlignInColumn:
                    var align = _settings.EndOfLineCommentColumnSpaceAlignment - code.Length;
                    return string.Format("{0}{1}{2}{3}", code, new string(' ', Math.Max(align - 1, 1)), EndOfLineComment,
                                                        _stupidLineEnding ? StupidLineEnding : string.Empty);
                default:
                    throw new InvalidEnumArgumentException();
            }
        }

        public override string ToString()
        {
            return Original;
        }

        private void AlignDims(int postition)
        {
            if (_segments[0].Trim().StartsWith("As "))
            {
                _segments[0] = string.Format("{0}{1}", new String(' ', _settings.AlignDimColumn - postition - 1), _segments[0].Trim());
                return;
            }
            var alignTokens = _segments[0].Split(new[] { " As " }, StringSplitOptions.None);
            if (alignTokens.Length <= 1)
            {
                return;
            }
            var gap = Math.Max(_settings.AlignDimColumn - postition - alignTokens[0].Length - 2, 0);
            _segments[0] = string.Format("{0}{1} As {2}", alignTokens[0].Trim(), new string(' ', gap),
                                         string.Join(" As ", alignTokens.Skip(1)));
        }
    }
}
