using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.SmartIndenter
{
    public class LogicalCodeLine
    {
        private readonly List<AbsoluteCodeLine> _lines = new List<AbsoluteCodeLine>();
        private AbsoluteCodeLine _rebuilt;
        private readonly IIndenterSettings _settings;

        public int IndentationLevel { get; set; }
        public bool AtProcedureStart { get; set; }
        public bool AtPropertyStart { get; set; }
        public string PropertyIdentifier { get; set; }
        public bool AtEnumTypeStart { get; set; }
        public bool InsideProcedureTypeOrEnum { get; set; }

        public bool IsEmpty
        {
            get { return _lines.All(x => x.IsEmpty); }
        }

        public LogicalCodeLine(AbsoluteCodeLine first, IIndenterSettings settings)
        {
            _lines.Add(first);
            _settings = settings;
        }

        public LogicalCodeLine(IEnumerable<AbsoluteCodeLine> lines, IIndenterSettings settings)
        {
            _lines = lines.ToList();
            _settings = settings;
        }

        public void AddContinuationLine(AbsoluteCodeLine line)
        {
            var last = _lines.Last();
            line.IsDeclarationContinuation = last.HasDeclarationContinuation && !line.ContainsOnlyComment;
            _lines.Add(line);
        }

        public int NextLineIndents
        {
            get
            {
                RebuildContinuedLine();
                if (_rebuilt.ContainsOnlyComment)
                {
                    return 0;
                }
                var indents = _rebuilt.Segments.Count() < 2
                    ? _rebuilt.NextLineIndents
                    : _rebuilt.Segments.Select(s => new AbsoluteCodeLine(s, _settings)).Select(a => a.NextLineIndents).Sum();

                if (_rebuilt.ContainsIfThenWithColonNoElse)
                {
                    indents--;
                }
                return indents;

            }
        }

        public int EnumTypeIndents
        {
            get
            {
                if (!IsEnumOrTypeMember)
                {
                    return 0;
                }
                return _settings.IndentEnumTypeAsProcedure &&
                       AtEnumTypeStart &&
                       IsCommentBlock && 
                       !_settings.IndentFirstCommentBlock
                    ? 0
                    : 1;
            }
        }

        public int Outdents
        {
            get
            {
                RebuildContinuedLine();
                if (_rebuilt.Segments.Count() < 2)
                {
                    return _rebuilt.Outdents;
                }
                var baseSegment = new AbsoluteCodeLine(_rebuilt.Segments.First(), _settings);
                return baseSegment.Outdents;
            }
        }

        private void RebuildContinuedLine()
        {
            if (_rebuilt != null) return;
            if (_lines.Count == 1)
            {
                _rebuilt = _lines.First();
                return;
            }
            var code = _lines.Aggregate(string.Empty, (c, line) => c + line.ContinuationRebuildText);
            _rebuilt = new AbsoluteCodeLine(code, _settings);
        }

        public bool IsProcedureStart
        {
            get { return _lines.Any(x => x.IsProcedureStart); }
        }

        public bool IsPropertyStart
        {
            get { return _lines.Any(x => x.IsPropertyStart); }
        }

        public bool IsProcedureEnd
        {
            get { return _lines.Any(x => x.IsProcedureEnd); }
        }

        public bool IsEnumOrTypeStart
        {
            get { return _lines.Sum(x => x.EnumOrTypeStarts) > _lines.Sum(x => x.EnumOrTypeEnds); }
        }

        public bool IsEnumOrTypeEnd
        {
            get { return _lines.Sum(x => x.EnumOrTypeEnds) > _lines.Sum(x => x.EnumOrTypeStarts); }
        }

        public bool IsEnumOrTypeMember { get; set; }

        public bool IsDeclaration
        {
            get { return _lines.All(x => x.IsDeclaration || x.IsDeclarationContinuation); }
        }

        public bool IsCommentBlock
        {
            get { return _lines.All(x => x.ContainsOnlyComment); }
        }

        private static readonly Regex OperatorIgnoreRegex = new Regex(@"^(\d*\s)?\s*[+&]\s", RegexOptions.IgnoreCase);

        public string Indented()
        {
            if (!_lines.Any())
            {
                return string.Empty;
            }

            if (_lines.Count == 1)
            {
                return _lines.First().Indent(IndentationLevel, AtProcedureStart);
            }

            if (!_settings.AlignContinuations || _lines.First().ContainsOnlyComment || IsEnumOrTypeMember)
            {
                return string.Join(Environment.NewLine, _lines.Select(line => line.Indent(IndentationLevel, AtProcedureStart)));
            }

            var output = new List<string>();
            var current = _lines.First().Indent(IndentationLevel, AtProcedureStart);
            var commentPos = string.IsNullOrEmpty(_lines.First().EndOfLineComment) ? 0 : current.Length - _lines.First().EndOfLineComment.Length;
            output.Add(current);
            var alignment = FunctionAlign(current, _lines[1].Escaped.Trim().StartsWith(":="));

            for (var i = 1; i < _lines.Count; i++)
            {
                var line = _lines[i];
                if (line.IsDeclarationContinuation && !line.IsProcedureStart)
                {
                    output.Add(line.Indent(IndentationLevel, AtProcedureStart));
                    continue;
                }
                if (line.ContainsOnlyComment)
                {
                    commentPos = alignment;
                }
                if (commentPos > 0)
                {
                    output.Add(line.Indent(commentPos, AtProcedureStart, true));
                    continue;
                }

                var operatorAdjust = _settings.IgnoreOperatorsInContinuations && OperatorIgnoreRegex.IsMatch(line.Original) ? 2 : 0;              
                current = line.Indent(Math.Max(alignment - operatorAdjust, 0), AtProcedureStart, true);
                output.Add(current);
                alignment = FunctionAlign(current, i + 1 < _lines.Count && _lines[i + 1].Escaped.Trim().StartsWith(":="));
                commentPos = string.IsNullOrEmpty(line.EndOfLineComment) ? 0 : current.Length - line.EndOfLineComment.Length;
            }

            return string.Join(Environment.NewLine, output);
        }

        public override string ToString()
        {
            return _lines.Aggregate(string.Empty, (x, y) => x + y.ToString());
        }

        private static readonly Regex StartIgnoreRegex = new Regex(@"^(\d*\s)?\s*[LR]?Set\s|^(\d*\s)?\s*Let\s|^(\d*\s)?\s*(Public\s|Private\s)?Declare\s(PtrSafe\s)?(Function|Sub)|^(\d*\s+)", RegexOptions.IgnoreCase);
        private readonly Stack<AlignmentToken> _alignment = new Stack<AlignmentToken>();
        private int _nestingDepth;

        //The splitNamed parameter is a straight up hack for fixing https://github.com/rubberduck-vba/Rubberduck/issues/2402
        private int FunctionAlign(string line, bool splitNamed)
        {
            line = new StringLiteralAndBracketEscaper(line).EscapedString;
            var stackPos = _alignment.Count;

            for (var index = StartIgnoreRegex.Match(line).Length + 1; index <= line.Length; index++)
            {
                var character = line.Substring(index - 1, 1);
                switch (character)
                {
                    case "\a":
                    case "\x2":
                        break;
                    case "(":
                        //Start of another function => remember this position
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Function, index, ++_nestingDepth));
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 1, _nestingDepth));
                        break;
                    case ")":
                        //Function finished => Remove back to the previous open bracket
                        while (_alignment.Any())
                        {
                            var finished = _alignment.Count == stackPos + 1;                            
                            var token =_alignment.Pop();
                            if (token.NestingDepth < _nestingDepth )
                            {
                                _alignment.Push(token);
                                finished = true;
                            }
                            if (finished)
                            {
                                _nestingDepth--;
                                break;
                            }
                        }
                        break;
                    case " ":
                        if (index + 3 < line.Length && line.Substring(index - 1, 3).Equals(" = "))
                        {
                            //Space before an = sign => remember it to align to later
                            if (!_alignment.Any(a => a.Type == AlignmentTokenType.Equals || a.Type == AlignmentTokenType.Variable))
                            {
                                _alignment.Push(new AlignmentToken(AlignmentTokenType.Equals, index + 2, _nestingDepth));
                            }
                        }
                        else if (!_alignment.Any() && index < line.Length - 2)
                        {
                            //Space after a name before the end of the line => remember it for later
                            _alignment.Push(new AlignmentToken(AlignmentTokenType.Variable, index, _nestingDepth));
                        }
                        else if (index > 5 && line.Substring(index - 6, 6).Equals(" Then "))
                        {
                            //Clear the collection if we find a Then in an If...Then and set the
                            //indenting to align with the bit after the "If "
                            while (_alignment.Count > 1)
                            {
                                _alignment.Pop();
                            }
                        }
                        break;
                    case ",":
                        //Start of a new parameter => remember it to align to
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 1, _nestingDepth));
                        break;
                    case ":":
                        if (line.Substring(index - 1, 2).Equals(":="))
                        {
                            //A named paremeter => remember to align to after the name
                            _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 3, _nestingDepth));
                        }
                        else if (line.Substring(index, 2).Equals(": "))
                        {
                            //A new line section, so clear the brackets
                            _alignment.Clear();
                            index++;
                        }
                        break;
                }
            }
            //If we end with a comma or a named parameter, get rid of all other comma alignments
            if (line.EndsWith(", _") || line.EndsWith(":= _") || splitNamed)
            {
                while (_alignment.Any() && _alignment.Peek().Type == AlignmentTokenType.Parameter)
                {
                    _alignment.Pop();
                }
            } 
            else if (line.EndsWith("( _"))   //If we end with a "( _", remove it and the space alignment after it'
            {
                _alignment.Pop();
                _alignment.Pop();
            }

            var output = 0;
            var fallback = 0;
            //Get the position of the unmatched bracket and align to that
            foreach (var align in _alignment.Reverse())
            {
                switch (align.Type)
                {
                    case AlignmentTokenType.Parameter:
                        output = align.Position - 1;
                        break;
                    case AlignmentTokenType.Function:
                    case AlignmentTokenType.Equals:
                        output = align.Position;
                        break;
                    default:
                        fallback = align.Position - 1;
                        break;
                }
            }

            if (fallback == 0 || fallback >= line.Length - 1)
            {
                fallback = !_alignment.Any() ? (_settings.IndentSpaces * 2) : 0;
            }
            else
            {
                fallback = fallback + 1;
            }

            return output > 0 ? output : fallback;
        }
    }
}
