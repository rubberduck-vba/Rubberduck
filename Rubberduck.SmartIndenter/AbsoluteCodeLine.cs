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
        private const string StringPlaceholder = "\a";
        private static readonly Regex StringLiteralRegex = new Regex("\"(?:[^\"]+|\"\")*\"");
        private static readonly Regex StringReplaceRegex = new Regex(StringPlaceholder);
        private static readonly Regex LineNumberRegex = new Regex(@"^(?<number>\d+)\s+(?<code>.*)", RegexOptions.ExplicitCapture);
        private static readonly Regex EndOfLineCommentRegex = new Regex(@"^(?!(Rem\s)|('))(?<code>.*)(\s(?<comment>'.*))$", RegexOptions.ExplicitCapture);
        private static readonly Regex ProcedureStartRegex = new Regex(@"^(Public\s|Private\s|Friend\s)?(Static\s)?(Sub|Function|Property\s(Let|Get|Set))");
        private static readonly Regex ProcedureStartIgnoreRegex = new Regex(@"^[LR]?Set\s|^Let\s|^(Public|Private)\sDeclare\s(Function|Sub)");
        private static readonly Regex ProcedureEndRegex = new Regex(@"^End\s(Sub|Function|Property)");
        private static readonly Regex TypeEnumStartRegex = new Regex(@"^(Public\s|Private\s)?(Enum\s|Type\s)");
        private static readonly Regex TypeEnumEndRegex = new Regex(@"^End\s(Enum|Type)");
        private static readonly Regex InProcedureInRegex = new Regex(@"^(Else)?If\s.*\sThen$|^(Else)?If\s.*\sThen\s.*\sElse$|^Else$|^Case\s|^With|^For\s|^Do$|^Do\s|^While$|^While\s|^Select Case");
        private static readonly Regex InProcedureOutRegex = new Regex(@"^(Else)?If\s.*\sThen\s.*(?<!\sElse)$|^Else$|ElseIf\s.*\sThen$|^Case\s|^End With|^Next\s|^Next$|^Loop$|^Loop\s|^Wend$|^End If$|^End Select");
        private static readonly Regex DeclarationRegex = new Regex(@"^(Dim|Const|Static|Public|Private)\s.*\sAs\s");
        private static readonly Regex PrecompilerInRegex = new Regex(@"^#(Else)?If\s.+Then$|^#Else$");
        private static readonly Regex PrecompilerOutRegex = new Regex(@"^#ElseIf\s.+Then|^#Else$|#End\sIf$");
        private static readonly Regex SingleLineIfRegex = new Regex(@"^If\s.*\sThen\s.*(?<!\sElse)$");
        private static readonly Regex SingleLineElseIfRegex = new Regex(@"^ElseIf\s.*\sThen\s.*(?<!\sElse)$");

        private readonly IIndenterSettings _settings;
        private uint _lineNumber;
        private bool _numbered;
        private string _code;
        private readonly string[] _segments;
        private List<string> _strings;

        public AbsoluteCodeLine(string code, IIndenterSettings settings)
        {
            _settings = settings;
            _code = code;
            Original = code;

            ExtractStringLiterals();
            ExtractLineNumber();
            ExtractEndOfLineComment();

            _segments = _code.Split(new[] { ": " }, StringSplitOptions.None);
        }

        private void ExtractStringLiterals()
        {
            _strings = new List<string>();
            var matches = StringLiteralRegex.Matches(_code);
            if (matches.Count == 0) return;
            foreach (var match in matches)
            {
                _strings.Add(match.ToString());
            }
            _code = StringLiteralRegex.Replace(_code, StringPlaceholder);
        }

        private void ExtractLineNumber()
        {
            var match = LineNumberRegex.Match(_code);
            if (match.Success)
            {
                _numbered = true;
                _lineNumber = Convert.ToUInt32(match.Groups["number"].Value);
                _code = match.Groups["code"].Value.Trim();
            }
            else
            {
                _code = _code.Trim();
            }
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

        public string Original { get; private set; }
        
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
            get { return !IsEmpty && DeclarationRegex.IsMatch(_code); }
        }

        public bool HasContinuation
        {
            get { return _code.EndsWith(" _") || EndOfLineComment.EndsWith(" _"); }
        }

        public bool IsPrecompilerDirective
        {
            get { return Original.TrimStart().StartsWith("#"); }
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
                outs -= _segments.Count(s => SingleLineIfRegex.IsMatch(s));

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
            if (IsEmpty || (ContainsOnlyComment && !_settings.AlignCommentsWithCode))
            {
                return Original;
            }

            if ((IsPrecompilerDirective && _settings.ForceCompilerDirectivesInColumn1) ||
                (IsBareDebugStatement && _settings.ForceDebugStatementsInColumn1) ||
                (atProcStart && !_settings.IndentFirstCommentBlock && ContainsOnlyComment) ||
                (atProcStart && !_settings.IndentFirstDeclarationBlock && IsDeclaration))
            {
                indents = 0;
            }

            var number = _numbered ? _lineNumber.ToString(CultureInfo.InvariantCulture) : string.Empty;
            var gap = Math.Max((absolute ? indents : _settings.IndentSpaces * indents) - number.Length, number.Length > 0 ? 1 : 0);
            AlignDims(gap);

            var code = string.Join(": ", _segments);
            if (_strings.Any())
            {
                code = _strings.Aggregate(code, (current, literal) => StringReplaceRegex.Replace(current, literal, 1));
            }

            code = string.Join(string.Empty, number, new string(' ', gap), code);
            if (string.IsNullOrEmpty(EndOfLineComment))
            {
                return code;
            }

            var position = Original.LastIndexOf(EndOfLineComment, StringComparison.Ordinal);
            switch (_settings.EndOfLineCommentStyle)
            {
                case EndOfLineCommentStyle.Absolute:
                    return string.Format("{0}{1}{2}", code, new string(' ', Math.Max(position - code.Length, 1)), EndOfLineComment);
                case EndOfLineCommentStyle.SameGap:
                    var uncommented = Original.Substring(0, position - 1);
                    return string.Format("{0}{1}{2}", code, new string(' ', uncommented.Length - uncommented.TrimEnd().Length + 1), EndOfLineComment);
                case EndOfLineCommentStyle.StandardGap:
                    return string.Format("{0}{1}{2}", code, new string(' ', _settings.IndentSpaces * 2), EndOfLineComment);
                case EndOfLineCommentStyle.AlignInColumn:
                    var align = _settings.EndOfLineCommentColumnSpaceAlignment - code.Length;
                    return string.Format("{0}{1}{2}", code, new string(' ', Math.Max(align - 1, 1)), EndOfLineComment);
                default:
                    throw new InvalidEnumArgumentException();
            }
        }

        private void AlignDims(int postition)
        {
            if (!DeclarationRegex.IsMatch(_segments[0]) || IsProcedureStart) return;
            var alignTokens = _segments[0].Split(new[] { " As " }, StringSplitOptions.None);
            var gap = Math.Max(_settings.AlignDimColumn - postition - alignTokens[0].Length - 2, 0);
            _segments[0] = string.Format("{0}{1} As {2}", alignTokens[0].Trim(),
                                                          (!_settings.AlignDims) ? string.Empty : new string(' ', gap),
                                                          string.Join(" As ", alignTokens.Skip(1)));
        }
    }
}
