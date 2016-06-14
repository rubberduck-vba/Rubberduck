﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace Rubberduck.SmartIndenter
{
    internal class LogicalCodeLine
    {
        private List<AbsoluteCodeLine> _lines = new List<AbsoluteCodeLine>();
        private AbsoluteCodeLine _rebuilt;
        private readonly IIndenterSettings _settings;

        public int IndentationLevel { get; set; }
        public bool AtProcedureStart { get; set; }

        public bool IsEmpty
        {
            get { return _lines.All(x => x.IsEmpty); }
        }

        public LogicalCodeLine(AbsoluteCodeLine first, IIndenterSettings settings)
        {
            _lines.Add(first);
            _settings = settings;
        }

        public void AddContinuationLine(AbsoluteCodeLine line)
        {
            _lines.Add(line);
        }

        public int NextLineIndents
        {
            get
            {
                RebuildContinuedLine();
                return _rebuilt.NextLineIndents;
            }
        }

        public int Outdents
        {
            get
            {
                RebuildContinuedLine();
                return _rebuilt.Outdents;
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
            get { return _lines.Any(x => x.IsProcedureStart) && !_lines.Any(x => x.IsProcedureEnd); }
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
            get { return _lines.All(x => x.IsDeclaration); }
        }

        public bool IsCommentBlock
        {
            get { return _lines.All(x => x.ContainsOnlyComment); }
        }

        private static readonly Regex OperatorIgnoreRegex = new Regex(@"^(\d*\s)?\s*[+&]\s");

        public string Indented()
        {
            if (_lines.Count <= 1)
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
            var alignment = FunctionAlign(current, true);

            foreach (var line in _lines.Skip(1))
            {
                if (commentPos > 0)
                {
                    output.Add(line.Indent(commentPos, AtProcedureStart, true));
                    continue;
                }

                var operatorAdjust = _settings.IgnoreOperatorsInContinuations && OperatorIgnoreRegex.IsMatch(line.Original) ? 2 : 0;
                current = line.Indent(Math.Max(alignment - operatorAdjust, 0), AtProcedureStart, true);
                output.Add(current);
                alignment = FunctionAlign(current, false);
                commentPos = string.IsNullOrEmpty(line.EndOfLineComment) ? 0 : current.Length - line.EndOfLineComment.Length;
            }

            return string.Join(Environment.NewLine, output);
        }

        private static readonly Regex StartIgnoreRegex = new Regex(@"^(\d*\s)?\s*[LR]?Set\s|^(\d*\s)?\s*Let\s|^(\d*\s)?\s*(Public|Private)\sDeclare\s(Function|Sub)|^(\d*\s+)");
        private readonly Stack<AlignmentToken> _alignment = new Stack<AlignmentToken>();

        private int FunctionAlign(string line, bool firstLine)
        {
            var stackPos = _alignment.Count;

            for (var index = StartIgnoreRegex.Match(line).Length + 1; index <= line.Length; index++)
            {
                var character = line.Substring(index - 1, 1);
                switch (character)
                {
                    case "\"":
                        //A String => jump to the end of it
                        while (!line.Substring(index++, 1).Equals("\"")) { }
                        break;
                    case "(":
                        //Start of another function => remember this position
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Function, index));
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 1));
                        break;
                    case ")":
                        //Function finished => Remove back to the previous open bracket
                        while (_alignment.Any())
                        {
                            var finished = _alignment.Count == stackPos + 1 || _alignment.Peek().Type == AlignmentTokenType.Function;
                            _alignment.Pop();
                            if (finished)
                            {
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
                                _alignment.Push(new AlignmentToken(AlignmentTokenType.Equals, index + 2));
                            }
                        }
                        else if (!_alignment.Any() && index < line.Length - 2)
                        {
                            //Space after a name before the end of the line => remember it for later
                            _alignment.Push(new AlignmentToken(AlignmentTokenType.Variable, index));
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
                        _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 1));
                        break;
                    case ":":
                        if (line.Substring(index - 1, 2).Equals(":="))
                        {
                            //A named paremeter => remember to align to after the name
                            _alignment.Push(new AlignmentToken(AlignmentTokenType.Parameter, index + 3));
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
            if (line.EndsWith(", _") || line.EndsWith(":= _"))
            {
                while (_alignment.Any() && _alignment.Peek().Type == AlignmentTokenType.Parameter)
                {
                    _alignment.Pop();
                }
            }

            //If we end with a "( _", remove it and the space alignment after it
            if (line.EndsWith("( _"))
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
                fallback = !_alignment.Any() && firstLine ? (_settings.IndentSpaces * 2) : 0;
            }
            else
            {
                fallback = fallback + 1;
            }

            return output > 0 ? output : fallback;
        }
    }
}
